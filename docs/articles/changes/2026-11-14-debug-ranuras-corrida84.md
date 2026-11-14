# Diagnóstico transitorio de ranuras — corrida 84

## Purpose

Identificar dónde difiere la corrida real de la simulación hasta 30/05/2027.
Instrumentación únicamente: sin cambiar restricciones, fechas, búsqueda ni datos.

## Evidencia leída

- Libro5.xlsx, Semestres, filas 21–68: CALI III-A/B y IV-A/B terminan el 23/07/2027. Última materia: 48 h, viernes 13:00–17:00. Ocho sesiones por grupo posteriores al 30/05: 128 h entre cuatro grupos.
- Log local del 14/11/2026: ampliaciones manuales antes de corrida 84; última duración registrada, 238 días. No registra fechas exactas de esa ventana efectiva. ModoVerboso=False; HardViolations=5. Esto no permite atribuir las cinco infracciones a CALI.
- Base temporal operativa, consulta de solo lectura: corrida 84, individuo óptimo 40. Aulas 36–39: materias 18/24 con patrón 5, 48 h y rango estimado 07/05–29/07/2027. El rango estimado termina después de la última sesión exportada.
- Configuración actual de esas aulas: ventana 1,2,3,4,5; 24 relaciones con OrigenPatron=LEGADO. No existe instantánea del origen del patrón de corrida 84: parametros_json={}. Los genes guardados sí confirman los días restringidos.

## Causa sustentada y límite

El evaluador y la fábrica liberan patrones AUTOMATICO; conservan LEGADO.
Dos materias de 48 h restringidas al viernes requieren 24 viernes por grupo.
Las ranuras sobrantes de lunes a jueves no satisfacen esa restricción.
LUN-VIE ya es terminal: la escalera no agrega otro día laborable.
La simulación con patrones libres no representa estos datos.

Además, el escalador evalúa contexto.Ventana, no el límite original supuesto.
Una ventana manual ampliada puede volver factible el plan y omitir la propuesta.
La corrida anterior no registra el motivo exacto del descarte; las nuevas trazas lo capturan.
No se han convertido patrones LEGADO a AUTOMATICO ni modificado datos.

## Changes

Marcador común: `[AG-RANURAS-TMP]`.

- AgFormGeneradorHorarios: operación, versión y activación del diagnóstico. Esta cabecera INFO permite distinguir diagnóstico apagado de ausencia del código nuevo.
- AgEvaluadorDemandaVentana: corrida/traza/aula/sede, plazo efectivo, demanda, capacidad, origen y días efectivos por materia, pendientes y proyección diferenciada.
- AgPlanificadorOleadas: comparación de referencia/candidato y razón de selección o error anterior al candidato. Parámetro opcional de correlación; contratos anteriores conservados en las llamadas fuente.
- AgPlanificadorRanuras: calendario, demanda, nodos/presupuesto, búsqueda agotada, bandas y resultado. Sin logs por nodo.
- AgFabricaIndividuos: construcción de caché por corrida/aula, demanda y patrones; origen de FechasGen y resultado de la primera semilla. Sin repetir genes de toda la población.
- AgEscaladorRanuraAula: plazo efectivo frente a plazo del aula; motivos de omisión, propuestas, coincidencias de huellas, idempotencia, commit o reversión. Sin contenidos de SQL ni huellas completas.
- AgMaterializadorSesiones y AgExpansorSesiones: composición antes/después del sellado, entrada con FechasGen o calendario legacy, salida, errores y resumen guardado.

## How to use

Tras autorizar compilación y cargar el binario actualizado:

1. Activar **Modo Diagnóstico (DEBUG)** antes del diagnóstico/pre-vuelo y de Optimizar. La casilla inicia desmarcada.
2. Registrar el plazo efectivo mostrado. Para contrastar el caso original, usar 30/05/2027 sin ampliar previamente. No modificar patrones durante la captura.
3. Revisar `%LOCALAPPDATA%/ConversorPDF/log/log_yyyyMMdd.txt`. Buscar `AG-RANURAS-TMP`; comprobar cabecera `activo=True`.
4. Seguir EVALUAR → BUSQUEDA → COMPARAR → ELEGIR → ESCALADOR_OMITIR o PROPUESTA → GUARDAR_HUELLAS/COMMIT → FABRICA_FECHAS/SEMILLA → STAGING/EXPANDIR.
5. Pre-vuelo puede tener corrida=0: correlacionar por traza y aula. CALI objetivo: aulas 36–39 en la base analizada. El ID de corrida nueva será distinto de 84.
6. Desactivar la casilla después de capturar. Logs DEBUG funcionan también en Release si el modo diagnóstico está activo. Retirar instrumentación buscando el marcador y los parámetros `diagnostico` añadidos cuando concluya la investigación.

## Test coverage

Sin nuevas reglas de negocio ni pruebas que repliquen mensajes de log.
Revisión estática de tipos, llamadas y diferencias; `git diff --check` sin errores.
No se compiló, ejecutó pruebas, instaló binario ni lanzó AG: pendiente de autorización.
No se confirma corrección funcional de la asignación con esta instrumentación.
