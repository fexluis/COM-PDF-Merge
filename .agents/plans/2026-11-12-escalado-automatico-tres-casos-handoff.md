# Handoff: Escalado Automático de Ventanas en Tres Casos (R-066)

**Fecha**: 2026-11-12  
**Plan asociado**: `.agents/plans/2026-11-12-escalado-automatico-tres-casos-plan.md`  
**Registro**: Fila `R-066` en `.agents/plans/REGISTRY.md` (Ronda 358–365, Fases 1734–1741)  
**Estado actual**: `closed`  
**Commit inicial de plan**: `0bcc7c1`  
**Commit final de pruebas**: `852708d`  
**Gate de cierre**: Full Verify PASS (1253/1253 tests PASS, 0 fallos, exit 0)  
**Evidencia final**: `_evidence/20261112-192936/test.trx`  
**Release y smoke interactivo**: fuera de alcance / pendientes de autorización de despliegue  

---

## 1. Contexto y Objetivos

El planificador académico y el algoritmo genético presentaban saturaciones silenciosas de carriles y desbordamiento de calendarios cuando la demanda total de sesiones cabía globalmente en la ventana nominal pero sobrecargaba días específicos (en especial Domingos en `VIE-DOM`, donde Viernes/Sábado aportan 3 ranuras y Domingo sólo 1 por semana).
Asimismo, los tres casos canónicos de ampliación de ventana requerían un tratamiento sistemático:
1. `LUN-JUE` con demanda 5W → escalado canónico a `LUN-VIE` ("1,2,3,4,5").
2. `MAR-VIE` con demanda 5W → escalado canónico a `LUN-VIE` ("1,2,3,4,5").
3. `VIE-DOM` con demanda 5W → escalado a `JUE-DOM` ("4,5,6,7"); y con demanda 6W → escalado sucesivo a `JUE-DOM` (mostrando demanda pendiente de 11 sesiones) y posteriormente a `MIE-DOM` ("3,4,5,6,7").

---

## 2. Resumen de Fases Ejecutadas (Fases 1734–1741)

- **T1 (Fase 1734)**: Contrato de aceptación y reproducción de cuellos de botella (`AgEscaladoTresCasosTests.cs`). Demostración de saturación asimétrica en Domingo y verificación de escalera.
- **T2 (Fase 1735)**: Origen explícito de patrones (`AgOrigenPatron`: `Automatico`, `Fijo`, `Legado`), columna en entidad `AgAulaAsignatura` y migración DDL `MigrationV80_AddOrigenPatronAgAulaAsignatura.cs` registrada en `SqliteHelper` y `ConversorPDF.XLL.csproj`.
- **T3 (Fase 1736)**: Partición de carriles y evaluación de demanda en `AgParticionCarriles.cs`. Permite evaluar alternativas cuando los patrones son automáticos y acepta mejoras temporales sin degradación de factibilidad.
- **T4 (Fase 1737)**: Evaluador de demanda compartida y decisión de escalado: creación de `AgResultadoEvaluacionVentana.cs` y `AgEvaluadorDemandaVentana.cs`, eliminando la salida prematura en `AgEscaladorVentanaAula.cs` para detectar saturación de carriles aunque la capacidad global nominal sea suficiente.
- **T5 (Fase 1738)**: Aplicación transaccional e invalidación de patrones en `AgEscaladorVentanaAula.AplicarAsync`: actualiza `ag_aula` y `ag_arbol_nodo` y resetea `dias_patron = ''` únicamente para `origen_patron = 'AUTOMATICO'`, preservando patrones fijos y legados.
- **T6 (Fase 1739)**: Integración y unificación UI en `AgFormGeneradorHorarios.cs`: pre-vuelo compartido entre `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync`, diálogo modal detallado con demanda pendiente y exclusión mutua de eventos.
- **T7 (Fase 1740)**: Suite de pruebas integral para la matriz de casos (`AgEscaladoTresCasosMatrizTests.cs`). Validación de LUN-JUE 5W, MAR-VIE 5W, VIE-DOM 5W, VIE-DOM 6W multi-paso con demanda pendiente, terminalidad de MIE-DOM y LUN-VIE, y diferenciación de capacidad para sábado 4h (68 ranuras) vs 8h (51 ranuras).
- **T8 (Fase 1741)**: Suite completa Full Verify ejecutada: 1253/1253 tests PASS, 0 fallos, 0 regresiones. Actualización de baseline de 1223 a 1253 tests.

---

## 3. Puntos de Contacto y Arquitectura

- **Capa Domain**: `AgOrigenPatron.cs` (enum), `AgAulaAsignatura.OrigenPatron` (string con valor por defecto "LEGADO").
- **Capa Application**: `AgEvaluadorDemandaVentana.cs`, `AgResultadoEvaluacionVentana.cs`, `AgParticionCarriles.cs`.
- **Capa Infrastructure (XLL)**: `AgEscaladorVentanaAula.cs`, `MigrationV80_AddOrigenPatronAgAulaAsignatura.cs`, `SqliteHelper.cs`.
- **Capa UI (XLL)**: `AgFormGeneradorHorarios.cs` (unificación de pre-vuelo y modal de escalado con soporte de demanda pendiente).
- **Capa Tests**: `AgEscaladoTresCasosTests.cs`, `AgOrigenPatronTests.cs`, `AgEvaluadorDemandaVentanaTests.cs`, `AgEscaladorVentanaAulaPersistenciaTests.cs`, `AgEscaladoTresCasosMatrizTests.cs`.
