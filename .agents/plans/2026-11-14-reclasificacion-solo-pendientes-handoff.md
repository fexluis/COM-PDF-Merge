# Handoff: Reclasificación Exclusiva de Materias con Sesiones Pendientes dentro del Plazo (R-077)

**Fecha**: 2026-11-14  
**Plan asociado**: `.agents/plans/2026-11-14-reclasificacion-solo-pendientes-plan.md`  
**Registro**: Fila `R-077` en `.agents/plans/REGISTRY.md` (Ronda 376–383, Fases 1788–1797)  
**Estado final**: `closed` (ejecución Swarm 100% completada y sellada)  
**Commit verificado**: `93137d3` (T10 Gate Full) · `0b5e906` (Cierre en REGISTRY.md)  
**Baseline**: 1356 tests PASS (0 fallos, 0 regresiones)  
**Workspace**: `C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`

---

## 1. Problema Resuelto y Diagnóstico Original

En la versión anterior de `ConversorPDF.XLL.Academic`:
1. `AgFormGeneradorHorarios` obtenía candidatos directamente desde `IdentificarPatronesRedistribuibles(contexto)` **antes** de evaluar el plan fechado, sin consultar si las materias realmente tenían déficit o sesiones pendientes en el plazo.
2. `IdentificarPatronesRedistribuibles` no verificaba el estado del aula (`a.Estado == AgEstadoAula.Disponible`), permitiendo que aulas con estado no disponible (como el caso del aula 12 con `Estado = 4 (Aprobado)`) ingresaran a la lista de candidatas.
3. Al seleccionar cualquier asignatura de dicha aula o con estado desactualizado, `EvaluarInterno` lanzaba un `ArgumentException("Reclasificación ajena o configuración modificada...")`, que se propagaba como `ExcepcionNoControlada` en la interfaz gráfica.
4. Materias que ya caben completamente dentro del plazo figuraban como candidatas a liberar capacidad innecesariamente.

---

## 2. Solución Arquitectónica Implementada

- **Elegibilidad canónica**: `AgEscaladorRanuraAula.EsAulaElegible(aula)` exige estrictamente `Id > 0`, `AulaRef` no vacía y `Estado == AgEstadoAula.Disponible` (1). Aulas en estado Aprobado (4), Inactiva, Mantenimiento o Reservada quedan totalmente excluidas de generar candidatas y de autorizaciones.
- **Candidatas fechadas y acotadas al déficit**: Una asignatura es candidata si y solo si:
  - Pertenece a un aula elegible (`Disponible`).
  - Su origen de patrón es `LEGADO` con días asignados.
  - Presenta sesiones pendientes positivas (`DemandaRequerida > SesionesUbicadas`) en la evaluación fechada del plazo aprobado.
  - Si el aula cabe completamente en plazo (`CabeEnPlazo == true`), devuelve 0 candidatas.
  - Materias completas NUNCA son candidatas a reclasificación, aun si su aula tiene otras asignaturas pendientes.
- **Validación tipada de selecciones**: `AgEscaladorRanuraAula.ValidarSelecciones(...)` retorna `AgResultadoValidacionSeleccion` discriminando causas (`AULA_NO_ELEGIBLE`, `CONFIGURACION_MODIFICADA`, etc.) y emite `WARN` en el log con IDs exactos.
- **Aislamiento SQL y Enriquecimiento de Decision**: `Decision` expone `IdAula` y en `AplicarDetalladoAsync` la actualización se acota estrictamente con `WHERE id_aula = @IdAula`.
- **Sobrecargas compatibles**: Se mantiene la sobrecarga `IdentificarPatronesRedistribuibles(contexto, aula, evaluacion = null)` para que cada etapa intermedia compile limpiamente.
- **Flujo UI ordenado y robusto**:
  1. Capturar contexto y plazo.
  2. Evaluar plan fechado primero.
  3. Extraer candidatas con déficit positivo.
  4. Si no hay candidatas, omitir `FormAutorizarReclasificacion`.
  5. Si hay candidatas, permitir selección individual desmarcada por defecto.
  6. Validar tipadamente la selección con mensaje claro al usuario ante rechazo.
  7. Simular en copia y avanzar a escalón solo si persiste déficit.
  8. Captura de errores de negocio protegida con `try-catch` sin propagar `ExcepcionNoControlada`.

---

## 3. Resumen de Olas Ejecutadas

| Wave | Ronda | Task | Fase | Enfoque | Estado | Evidencia | Commit |
|---|---|---|---|---|---|---|---|
| 1 | 376 | T1 | 1788 | Modelo `AgReclasificacionPatron` con métricas de déficit | `done` | `scripts/verify.ps1 -Mode Build` | `929c61c` |
| 1 | 376 | T2 | 1789 | Desglose por asignatura en evaluador y tests unitarios | `done` | `_evidence/20261114-183431` | `a6c0acd` |
| 2 | 377 | T3 | 1790 | Calibración preventiva de fixtures en pruebas de redistribución | `done` | `_evidence/20261114-184000` | `d88714e` |
| 2 | 377 | T4 | 1791 | Diálogo UI `FormAutorizarReclasificacion` con metadatos y déficit | `done` | `_evidence/20261114-183727` | `8c5f211` |
| 3 | 378 | T5 | 1792 | Elegibilidad canónica, candidatas fechadas e IdAula en Decision / SQL | `done` | `_evidence/20261114-185016` | `5c356b6` |
| 4 | 379 | T6 | 1793 | Validación tipada de selección y manejo de rechazos controlados | `done` | `_evidence/20261114-185823` | `9689995` |
| 5 | 380 | T7 | 1794 | Pruebas unitarias de elegibilidad y rechazos controlados (28 tests) | `done` | `_evidence/20261114-190300` | `8df072d` |
| 6 | 381 | T8 | 1795 | Orquestación del flujo UI y captura controlada en generador | `done` | `_evidence/20261114-191000` | `151ccaa` |
| 7 | 382 | T9 | 1796 | Pruebas de aceptación end-to-end (20 tests) | `done` | `_evidence/20261114-191600` | `56774e0` |
| 8 | 383 | T10 | 1797 | Certificación Gate Full (1356 tests PASS) y calibración de baseline | `done` | `_evidence/20261114-192003` | `93137d3` |
| — | — | — | — | Sellado oficial de R-077 en `REGISTRY.md` | `closed` | `scripts/verify.ps1 -Mode Full` | `0b5e906` |
| — | — | — | — | Correcciones pre-release [P1, P1, P2] | `done` | `_evidence/20261114-195640` | `10a3802` |
| — | — | — | — | Corrección [P2 · T3, P2 · T4]: Diagnóstico sellado y resultado tipado | `done` | `_evidence/20261114-202146` | `e90d5af` |

---

## 4. Correcciones Pre-Release Verificadas

Se aplicaron las correcciones identificadas previas al release:
1. **[P1] Inclusión de `IdAula` en Huella de Acción e Invariante de Escritura**:
   - `AgEscaladorRanuraAula.HuellaAccion` migrado a `"propuesta-escalado-v4"` incluyendo `d.IdAula.ToString()`.
   - `AplicarDetalladoAsync` valida `d.IdAula == aulas[0].Id` y `rec.IdAula == idAula`, lanzando `InvalidOperationException` y abortando toda escritura si difiere.
   - Nueva prueba unitaria: `Aplicar_WhenIdAulaManipulado_ShouldRechazarYNoEscribir`.
2. **[P1] UI Bloqueante ante Rechazo o ArgumentException**:
   - `AgFormGeneradorHorarios.EvaluarYAplicarEscaladoInternoAsync` retorna `AgResultadoEscaladoUI.FalloAlAplicar` ante fallo de validación o excepción de argumentos.
   - Detiene de forma preventiva el avance del pipeline genético hacia pre-vuelo y optimización, exigiendo recalcular.
3. **[P2] Validación Estricta contra Déficit en `ValidarSelecciones`**:
   - `ValidarSelecciones` valida que toda materia seleccionada pertenezca a las candidatas con déficit positivo (`sesionesPendientes > 0`).
   - Retorna `AgResultadoValidacionSeleccion.Rechazo("MATERIA_SIN_DEFICIT", ...)` evitando que llamadas directas autoricen materias ya completas.
   - Nueva prueba unitaria: `ValidarSelecciones_WhenMateriaCompletaSinDeficit_ShouldRetornarRechazoMateriaSinDeficit`.
4. **[P2 · T3] Reutilización de Evaluación Inicial y Diagnóstico Sellado**:
   - `AgFormGeneradorHorarios` obtiene candidatas directamente desde `decisionesEscalado.ReclasificacionesPropuestas` de la evaluación inicial, eliminando la re-evaluación redundante de aulas.
   - `ValidarSelecciones` valida en memoria contra las candidatas selladas sin recalcular.
   - Simulación recalculada únicamente cuando existen cambios autorizados; de lo contrario reutiliza la evaluación inicial.
   - Se propaga cancelación `CancellationToken` en todas las rutas de evaluación y filtrado.
5. **[P2 · T4] Eliminación de Excepciones para Rechazos Esperados (Resultado Tipado)**:
   - `AgEscaladorRanuraAula.Evaluar` devuelve `AgResultadoEvaluacionEscalado : IReadOnlyList<Decision>`, encapsulando `.EsValido`, `.Motivo`, `.Mensaje`, `.IdAulaRechazada`, `.IdAsignaturaRechazada`.
   - `EvaluarInterno` retorna `AgResultadoEvaluacionEscalado.Rechazo(validacion)` sin lanzar `ArgumentException`.
   - La UI propaga motivo e identidades rechazadas al usuario.
   - Pruebas unitarias adaptadas para verificar el resultado tipado sin excepciones de flujo.

**Certificación de la suite**: **1359/1359 PASS** (0 fallos, 0 regresiones) con baseline calibrado.
