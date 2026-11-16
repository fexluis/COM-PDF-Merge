# Resolver cobertura de aulas aprobadas sin ampliar el calendario

Estado: PROPUESTO. Este documento autoriza y describe planificación, no ejecución.
Objetivo: cubrir todas las relaciones seleccionadas hasta el **30/05/2027**, ofreciendo el siguiente escalón a las aulas aprobadas con déficit verificable y preservando su horario productivo hasta validar y aprobar el reemplazo.

Referencia funcional: [plan detallado de cobertura e I5](2026-11-15-cobertura-aulas-aprobadas-i5-plan.md). Esta hoja de ruta organiza su ejecución; no elimina requisitos del plan detallado. Ante contradicciones, resolverlas antes de implementar.

## Reglas que no se negocian

1. La aprobación protege la escritura; no impide evaluar ni presentar propuestas.
2. Un plan aprobado completo y verificable se conserva, aunque una simulación desde cero no encuentre solución.
3. Déficit verificable con escalón disponible: ofrecerlo antes de devolver I5. CALI I-B debe poder evaluar **MAR-VIE → LUN-VIE**.
4. Una confirmación autoriza sólo el paso, aula, alcance y revisión mostrados. Redistribución LEGADO requiere consentimiento por asignatura; FIJO no se convierte por consentimiento genérico.
5. Un paso por aula y corrida/intento. Un reintento no consume otro paso; la reevaluación no reabre el formulario en esa corrida.
6. No ampliar la fecha final ni ocultar relaciones seleccionadas para obtener éxito.
7. No borrar producción, histórico o auditoría para sortear I5. No convertir masivamente aulas a Disponible.
8. No compilar ni ejecutar pruebas sin aprobación explícita. No publicar ni liberar con verificaciones pendientes.

## Evidencia de partida

| Hallazgo | Evidencia investigada |
|---|---|
| #95 abortó antes de materializar | Log del 15/11/2026, 20:43:32: I5 para `3\|1\|1\|I\|B\|PRESENCIAL`, duración 31 ms. |
| Antes del bloqueo se eliminó #94 | Mismo log: purga de 184 vinculaciones y 2064 sesiones. |
| El escalador excluye aprobadas | `AgEscaladorRanuraAula.cs:118`; el filtro se reutiliza en evaluación y validación/aplicación. |
| La fábrica omite aprobadas | `AgFabricaIndividuos.cs:73`. |
| I5 no certifica el plan productivo | `AgGeneradorHorarios.ValidarAprobadas` simula demanda desde cero. |
| La compuerta puede aceptar omisiones | Compara genes existentes con staging, no demanda seleccionada con solución completa. |

Las líneas son referencias de la investigación, no ubicaciones estables. La ausencia de sesiones en staging/Excel no prueba ausencia en producción. Los valores de 256 h, 64 sesiones y déficit de tres sesiones del caso CALI deben reconciliarse contra el snapshot y los datos canónicos, no asumirse como resultado de esta investigación.

## Orden de trabajo

**T0 → T1 → T2 → T3 → T4 → T5 → T6 → T7.**

Cada etapa incorpora sus pruebas y documentación. No avanzar sobre contratos no resueltos ni marcar verificada una etapa cuyos controles no se ejecutaron. Separar cambios en unidades revisables; conservar los cambios ajenos existentes en evaluador, planificadores, proyecto y archivos de prioridad de horas.

### T0. Identificar origen canónico y cerrar contratos

**Trabajo de lectura, sin modificar bases ni ejecutar el generador.**

- Identificar tablas, estados y vínculos que acreditan un horario aprobado; incluir horarios manuales y aprobaciones históricas sin auditoría de promoción.
- Definir identidad completa de relación, sesión y revisión; distinguir ausencia verificable, origen ambiguo y fallo técnico.
- Determinar el mecanismo existente de permisos y quién puede confirmar la revisión y aprobar el reemplazo. Si no existe política, resolverla con el responsable funcional antes de implementar autorización.
- Recomendación para este flujo: reemplazo completo del aula confirmada, con todas sus relaciones enumeradas. Confirmar esta unidad antes de implementar; no mezclar reemplazo por aula y por asignatura implícitamente.
- Definir configuración candidata separada del horario productivo: persistencia, recuperación y abandono de pasos confirmados todavía insuficientes.

**Salida:** mapa de origen, contratos de lectura y autorización, unidad de reemplazo y ciclo de vida candidato documentados. No inventar SQL ni revisiones donde el origen no permita verificarlas.

### T1. Proteger histórico, intento y estados terminales

- Retirar la purga global del inicio; una corrida nueva sólo escribe y limpia sus propios resultados parciales.
- Conservar auditoría de promoción y referencias necesarias para reconstrucción histórica.
- Crear identidad estable de intento antes de la preparación, aun cuando no se cree `ag_corrida`.
- Registrar bloqueos, errores y cancelaciones con códigos estructurados; cerrar corridas creadas sin dejarlas EnProgreso.
- Establecer estado visual terminal antes de modales/retornos y rechazar callbacks tardíos por identidad de ejecución.
- Impedir ejecuciones simultáneas. No desplazar trabajo a otro hilo hasta verificar que contexto y dependencias no usan controles ni Excel/COM.

**Aceptación:** repetir el escenario de #95 no elimina resultados anteriores y nunca deja la pantalla en Optimizando después del bloqueo.

### T2. Construir manifiesto y verificar aprobadas reales

- Crear manifiesto desde selección confirmada y relaciones requeridas, nunca desde genes o staging.
- Incluir clave completa, demanda original/efectiva, ventana y versión; persistir snapshot reproducible.
- Implementar lectura canónica mediante interfaz de Domain y repositorio de Infrastructure, con paginación completa y sin recortar sesiones exteriores a la ventana.
- Verificar identidad, horas, fechas, duración, festivos, duplicados y conflictos del plan aprobado.
- Asignar destino definitivo: conservar o replanificar; bloquear explícitamente estados no programables y orígenes no verificables.
- Error de lectura mantiene código técnico. Lectura exitosa con origen no verificable devuelve PLAN_NO_VERIFICABLE, no una propuesta sobre datos desconocidos.

**Aceptación:** I-A e I-B aparecen en el manifiesto con todas sus relaciones; un plan completo se conserva con sus fechas y docentes originales.

### T3. Simular y presentar un único escalón

- Separar `PuedeEvaluarPropuesta` de `PuedeAplicarPropuesta`; no ampliar globalmente el helper que habilita escritura.
- Preparación común para diagnóstico y optimización; evaluar todas las aulas afectadas, no sólo la primera que falla.
- Para déficit verificable, simular el siguiente escalón sin persistir: CALI I-B, MAR-VIE → LUN-VIE, añadiendo lunes y manteniendo 30/05/2027.
- Mostrar días añadidos, fechas recuperadas, déficit antes/después, patrones afectados y resultado sin/con redistribución LEGADO.
- Distinguir capacidad de calendario de factibilidad docente. No prometer sesiones recuperadas sin evidencias de planificación fechada; indicar restricciones pendientes.
- Respetar origen y límite de las escaleras canónicas e históricas del plan detallado, incluida semántica sábado M/T.

**Aceptación:** Aprobado recibe propuesta antes de I5; simular o cancelar no modifica configuración, permisos ni producción.

### T4. Confirmar con autorización revisionada e idempotencia

- Un formulario confirma el paso y autoriza explícitamente revisar el horario aprobado del alcance enumerado.
- Registrar actor, permiso, motivo, revisión de origen, ventana, alcance, patrones consentidos, origen/índice de escalera e identidad de propuesta/intento.
- Diferenciar huella previa, identidad de operación y huella del efecto aplicado; la operación no debe invalidar su propia autorización al modificar configuración.
- Verificar revisión y permisos; persistir configuración candidata y autorización en una transacción.
- Doble clic/reintento recupera el efecto aplicado. Cambios posteriores no autorizados invalidan el permiso.
- Reconstruir contexto y reevaluar una vez. Si el paso mejora pero no completa, permitir confirmación con advertencia explícita, persistirlo y detener por I5; siguiente oferta sólo en otra corrida.

**Aceptación:** confirmación repetida no agrega pasos ni autorizaciones; huella obsoleta no deja persistencia parcial; rechazo conserva el horario anterior.

### T5. Generar sin omisiones y validar cobertura completa

- Fábrica, cruce, mutación y coherencia sólo admiten aulas Replanificar con autorización vigente; desactivar fallback a aula nula cuando hay manifiesto explícito.
- Incorporar conservadas como sesiones inmutables y ocupaciones de aula/docente por intervalos reales.
- Excluir de la solución candidata las sesiones anteriores que el alcance autorizado reemplaza, sin retirarlas de producción ni dejar de bloquear otras ocupaciones.
- Mantener validación exacta genes ↔ staging nuevo y añadir manifiesto ↔ unión disjunta de conservadas y nuevas.
- Rechazar déficit, exceso, duplicación, relaciones inesperadas, fechas exteriores y conflictos. VACANTE no elimina demanda ni equivale a cobertura docente válida.
- Admitir Sólo conservación sin población artificial ni staging nuevo obligatorio.

**Aceptación:** quitar una relación completa falla aunque genes y staging coincidan; ninguna sesión supera 30/05/2027; se detectan cruces con conservadas.

### T6. Aprobar reemplazo, promover y mostrar resultado integral

- Mantener el horario anterior durante generación y validación del candidato; confirmar el reemplazo según la política resuelta en T0.
- Dentro de la transacción de promoción, releer/revalidar revisiones, autorización, cobertura y ocupaciones con control de concurrencia efectivo. Una huella calculada fuera de la transacción no basta.
- Sustituir sólo sesiones del alcance autorizado con vínculo a la revisión anterior; no depender exclusivamente del ID de staging para identificar qué se reemplaza.
- No tocar sesiones conservadas. Rollback completo ante fallo o cambio concurrente.
- Resultados y Excel incluyen nuevas y conservadas con origen visible, sin duplicar staging o producción.
- Resumen reconcilia aulas, relaciones, horas esperadas/cubiertas, vacantes y bloqueos con el manifiesto persistido.

**Aceptación:** una corrida nueva reemplaza en vez de duplicar; fallo parcial conserva producción anterior; exportación reconcilia cada relación.

### T7. Verificación integral y cierre

**Sólo ejecutar tras aprobación.**

1. Pruebas en memoria: conservación, propuestas, consentimiento, escaleras, idempotencia, demanda y cobertura.
2. Integración sobre SQLite temporal: lectura canónica, persistencia atómica, conservación del histórico, promoción, rollback y cambios concurrentes.
3. Caso CALI reproducible con festivos y disponibilidad docente explícita: I-A/I-B completas por relación hasta 30/05/2027; incluir un caso sin solución y verificar rechazo honesto.
4. UI: Failure, excepción, cancelación, respuesta durante cálculo y callbacks tardíos. Antes de usar Task.Run, verificar aislamiento de controles/COM y propagación de cancelación.
5. Exportación: comparar relaciones y horas contra el manifiesto, no sólo cantidad de filas o fecha máxima.
6. Regresión: Disponible, aprobadas completas, Sólo conservación, todas las escaleras y patrones LEGADO/FIJO.
7. Ejecutar controles focalizados, suite aplicable y luego Release/smoke. Verificación independiente de permisos, datos productivos y concurrencia; revisión nativa según el interruptor y consentimiento vigentes.

Cada rechazo debe comprobar código, diagnóstico y ausencia de modificaciones no autorizadas. Registrar comprobaciones fallidas, omitidas y pendientes; la compilación sola no acredita cierre.

## Resultados terminales esperados

| Situación | Resultado |
|---|---|
| Aprobada completa y verificable | Conservar y contar cobertura. |
| Déficit verificable y siguiente escalón disponible | Propuesta antes de I5. |
| Confirmación y cobertura completa | Continuar generación/validación; no promover automáticamente. |
| Rechazo explícito | I5 / REVISION_RECHAZADA, sin cambios. |
| Paso confirmado todavía insuficiente o escalera agotada | I5 / COBERTURA_INSUFICIENTE, con déficit residual y siguiente acción. |
| Origen aprobado no verificable tras lectura exitosa | I5 / PLAN_NO_VERIFICABLE. |
| Fallo técnico, cancelación o revisión obsoleta | Código específico; no disfrazarlo de déficit. |

Definir en T0 el código y mensaje para Disponible insuficiente: no mostrarle un diagnóstico que afirme que el aula está aprobada.

## Criterio de cierre

- [ ] Propuesta para aprobadas antes de I5 cuando corresponde; sin escritura previa al consentimiento.
- [ ] Un único paso confirmado por intento y sin ampliación de fecha final.
- [ ] Todas las relaciones seleccionadas cubiertas o bloqueo explícito; ninguna omisión silenciosa.
- [ ] Producción anterior e histórico preservados hasta reemplazo autorizado y transaccional.
- [ ] Cobertura, conflictos, permisos y revisiones comprobados sobre el conjunto completo.
- [ ] UI y corrida terminan en estado coherente; resultados y Excel reflejan el manifiesto.
- [ ] Evidencia funcional y verificaciones aplicables completas antes de release.

**Próximo paso:** iniciar T0 como investigación de sólo lectura; resolver sus contratos antes de implementar. La ejecución de pruebas/build sigue pendiente de aprobación explícita.
