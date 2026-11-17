# Extras conservadas en aulas aprobadas (política B)

## Decisión

Cambio de requisito decidido por el humano (verbatim): «B: extras como ocupación fija».

Las relaciones productivas de un aula aprobada que **no** forman parte del manifiesto requerido dejan de
bloquear la preparación. Se declaran como **no requeridas (extras)**, se conservan inmutables, cuentan como
**ocupación fija** para la detección de conflictos de aula y docente (no hay doble reserva), **no** añaden
demanda requerida y **no** se replantean. La producción no se modifica.

## Alcance implementado

1. **Manifiesto** (`AgManifiestoCobertura`): nueva colección explícita `RelacionesExtra` con parámetro de
   constructor opcional (los llamadores existentes siguen compilando) y helper `EsExtra(...)`. Las extras
   entran en la huella del manifiesto, de modo que cualquier cambio invalida autorizaciones obsoletas. No
   participan en la demanda requerida ni en `DetectarOmisiones`. Una extra que duplique una requerida se
   rechaza con `EXTRAS_INVALIDAS`.
2. **Constructor del manifiesto** (`AgConstructorManifiestoCobertura`): declara como extras las relaciones
   `Requerida != 1` de las aulas seleccionadas, con la misma validación de identidad que las requeridas.
   El lector además descubre materias exclusivamente productivas cuando la identidad completa determina
   una única aula seleccionada; las liga al manifiesto/huella antes de ratificar, sin fila AG auxiliar ni
   escritura de DB. Una identidad incompleta o ambigua sigue bloqueando.
3. **Lector/snapshot** (`AgPlanAprobadoRepository`, `AgPlanAprobadoSnapshot`): una relación productiva que
   coincide exactamente con una extra declarada se clasifica como extra y deja de emitir `RELACION_EXTRA`.
   `EstadoOrigen` sigue siendo `PLAN_NO_VERIFICABLE` para incertidumbre real: identidad ambigua o incompleta,
   múltiples vinculaciones, requeridas ausentes, subunidad desconocida, sesiones huérfanas o no verificables y
   relaciones productivas no reconciliables. Las extras se exponen en `RelacionesExtra` y entran en
   `HuellaOrigen`.
4. **Validador de cobertura** (`AgValidadorCoberturaSeleccion`): acepta sesiones de una extra declarada
   **sólo** con origen `CONSERVADA`; una extra como `NUEVA` o una relación no declarada siguen devolviendo
   `RELACION_INESPERADA`. Las extras no suman horas requeridas, ni déficit, ni cobertura, ni exceso de una
   requerida, y sí participan en la detección de `CRUCE_AULA`/`CRUCE_DOCENTE`.
5. **Validador de plan** (`AgValidadorPlanAprobado`): las extras no distorsionan horas esperadas, déficit ni
   exceso de las requeridas. Sus sesiones reciben la misma validación de identidad, duración, fecha,
   festivos, docente y duplicados. Los cruces extra-extra entran en diagnósticos globales antes de ratificar.
6. **Proyector** (`AgProyectorSolucionCobertura`): proyecta las extras de las aulas aprobadas como sesiones
   `CONSERVADA` inmutables para bloquear conflictos; nunca las inserta en staging ni en producción.
7. **UI** (`AgFormGeneradorHorarios`): el resumen/consentimiento por aula informa cuántas extras se conservan
   y que no se replantean; el bloqueo `PLAN_NO_VERIFICABLE` desaparece cuando el único motivo eran extras y
   su mensaje ya no las menciona. Se mantienen los guards T4/T6: sin ratificación vigente no hay generación ni
   promoción para aulas aprobadas.

## Invariantes preservados

- El límite temporal es la **ventana semestral seleccionada**; no se reintrodujo ninguna fecha fija
  (30/05/2027). `grep` sobre producción no encuentra `new DateTime(2027, 5, 30)` ni
  `VENTANA_EXCEDE_LIMITE_30_05_2027`.
- No se debilitó ni eliminó ninguna aserción existente. Las pruebas que describían el bloqueo por extras se
  conservan para relaciones **no declaradas** (`RELACION_EXTRA`, `RELACION_INESPERADA`), que siguen bloqueando.
- No se publicó, instaló, migró ni promocionó nada; no se modificó producción.

## Pruebas

- `AgManifiestoCoberturaTests`: extras separadas de la demanda requerida, no son omisión, la huella cambia al
  cambiar las extras y una extra que duplica una requerida se rechaza.
- `AgValidadorCoberturaSeleccionTests`: extra conservada no bloquea ni suma demanda; extra no cubre la
  requerida; extra como `NUEVA` es `RELACION_INESPERADA`; extra conservada bloquea cruce de aula y de docente;
  el proyector conserva la extra inmutable.
- `AgValidadorPlanAprobadoTests`: extras no distorsionan horas/déficit/exceso; extras ocupan aula y docente;
  extras no ocultan ambigüedad ni requerida ausente.
- `AgPlanAprobadoRepositoryTests`: extra declarada deja el plan en `REQUIERE_RATIFICACION` sin diagnósticos y
  se conserva; una extra declarada no oculta ambigüedad ni requerida ausente (`PLAN_NO_VERIFICABLE`).

El reporte anterior señaló 1598/1598 y build Debug sin artefactos de prueba localizables. En esta
continuación se observaron RED de colisión extra-extra, extra temporal/identidad inválida y materia
exclusivamente productiva: 3 fallos / 39 pruebas; `ConversorPDF.Tests/TestResults/politica-b-red.trx`.

**Límite separado T5/T6:** los controles duros rechazan colisiones con extras, pero la fábrica no reserva
proactivamente ranuras para ellas; la calidad de la búsqueda no está garantizada. No se implementan
eliminación/reemplazo del plan anterior ni rehidratación (T6).
