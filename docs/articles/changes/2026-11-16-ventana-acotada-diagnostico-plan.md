# Ventana acotada al período y diagnóstico real del plan canónico

## Decisión

Dos correcciones derivadas del análisis del log `log_20261116.txt`:

1. **El límite ya no es una fecha fija.** El tope es `periodos.periodo_inicio` → `periodos.periodo_fin`
   (para 2027-01: 2027-02-05 → 2027-05-30, 115 días). Una ventana manual que exceda ese rango se rechaza
   con `VENTANA_FUERA_DE_PERIODO`; no se acota en silencio ni se amplía la cobertura temporal.
2. **El terminal dice el motivo real.** `aula=N:PLAN_NO_VERIFICABLE:<motivo>` informa
   `RELACIONES_AUSENTES=ausentes/requeridas` y los códigos concretos del lector/validador, en lugar de una
   etiqueta genérica («origen ambiguo o relación extra») que no distinguía extras, ambigüedad, partición ni
   ausencia total.

Contexto productivo observado: las aulas **12, 13 y 14** (período 2027-01, estado Aprobado) se rechazaban
con `REVISION_RECHAZADA`. El lector canónico no encontraba **ninguna** vinculación ni sesión para 2027-01
(226 vinculaciones y 2255 sesiones existen solo para 2026-02), por lo que las 15 relaciones requeridas
(3+6+6) quedaban ausentes: `PLAN_NO_VERIFICABLE` por `RELACIONES_AUSENTES`, no por extras ni ambigüedad.

## Alcance implementado

1. **Resolutor de ventana** (`AgResolutorVentanaSemestre`): cuando el período tiene fechas canónicas y el
   usuario aporta un rango manual, se valida contención (`manualIni >= periodo_inicio` y
   `manualFin <= periodo_fin`). Fuera de rango devuelve
   `VENTANA_FUERA_DE_PERIODO:<manual>;periodo=<periodo>` sin crear `VentanaSemestre`. Las prioridades 2
   (festivos `PRESENCIAL`) y 3 (manual sin fechas de período) mantienen su comportamiento: allí no existe
   referente canónico que acotar. El origen sigue siendo `Manual`, de modo que la UI y las pruebas previas
   no cambian de semántica.
2. **Descriptor de diagnóstico** (`AgDescriptorDiagnosticoPlan`, nuevo): función pura que resume un
   `AgValidacionAulaPlan` en `RELACIONES_AUSENTES=n/m` más hasta `maxCodigos` diagnósticos ordenados
   ordinalmente, con `(+N)` cuando se trunca. Incluye los diagnósticos globales del lector porque también
   dejan el aula en `PlanNoVerificable`. Si no hay diagnóstico tipado devuelve `SIN_DIAGNOSTICO_CONOCIDO`:
   nunca inventa una causa.
3. **Terminal de preparación** (`AgFormGeneradorHorarios`): el bloqueo por aula usa el descriptor en lugar
   del texto fijo. Un período sin plan produce
   `aula=12:PLAN_NO_VERIFICABLE:RELACIONES_AUSENTES=3/3;RELACION_AUSENTE:…`.
4. **Log durable** (`AgFormGeneradorHorarios.Log`): además del panel `txtLogAG`, cada línea se emite con
   `LoggerHelper.Info`. `LoggerHelper.Minimo` es `NivelLog.Info`, por lo que `Debug` se descartaría: era la
   razón por la que `[PLAN CANÓNICO]` no aparecía en `log_*.txt` y el detalle se perdía al cerrar el
   formulario.
5. **Volumen controlado**: el resumen por aula se registra siempre; el detalle por relación solo cuando hay
   diagnóstico, déficit o exceso (`> 0.001`). Las extras se registran siempre por ser pocas.

## Invariantes preservados

- `PLAN_NO_VERIFICABLE` sigue siendo terminal: no se simula, no se escribe y no se ratifica sobre un origen
  desconocido.
- Ningún cambio convierte el aula en Disponible, oculta aulas ni auto-ratifica. T4 sigue siendo obligatorio
  antes de generar o promocionar.
- No se modificó el manifiesto ni la demanda: la ausencia se informa, no se rellena.
- Sin migración de esquema ni escritura sobre bases reales en esta unidad.

## Pruebas

- `AgResolutorVentanaSemestreTests`: ventana manual antes/después del período → `VENTANA_FUERA_DE_PERIODO`;
  ventana igual al período (115 días) → `Manual` válida; se conservan las pruebas previas de prioridades 2/3
  y del rango manual dentro del período.
- `AgDescriptorDiagnosticoPlanTests` (nuevo): conteo y código concreto, causa desconocida sin invención,
  inclusión de diagnósticos globales, truncado con `(+N)`, determinismo y argumento nulo.
- `AgPlanAprobadoRepositoryTests`: período sin filas productivas → todas las requeridas ausentes y
  `EstadoOrigen = PLAN_NO_VERIFICABLE`, sin diagnósticos inventados.
- `AgValidadorPlanAprobadoTests`: ese escenario deja el aula en `PlanNoVerificable`, emite
  `RELACION_AUSENTE:` por relación y el descriptor informa `RELACIONES_AUSENTES=2/2`.
- `AgValidadorCoberturaSeleccionTests`: se conserva `VentanaHastaJunio_NoEsBloqueadaPorUnaFechaFija`.

## Límites explícitos

- Esta unidad **no** desbloquea 12/13/14: mientras 2027-01 no tenga plan productivo, el rechazo es correcto.
  Cargar o aprobar ese plan es una decisión de datos, no de código, y conduce a `REQUIERE_RATIFICACION`.
- El descriptor resume; no sustituye la revisión de datos ni la ratificación auditada.
- `LoggerHelper` usa cola asíncrona y escritura best-effort: no garantiza entrega exacta ante cierre
  abrupto. Es el canal durable existente, no un almacén de autorización.
- El binario instalado el 2026-11-16 (18:19) es anterior a estos cambios y todavía contiene
  `VENTANA_EXCEDE_LIMITE_30_05_2027`; requiere recompilar y redesplegar para que el límite fijo desaparezca.
