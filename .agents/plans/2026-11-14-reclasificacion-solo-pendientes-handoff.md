# Handoff: Reclasificación Exclusiva de Materias con Sesiones Pendientes dentro del Plazo (R-077)

**Fecha**: 2026-11-14  
**Plan asociado**: `.agents/plans/2026-11-14-reclasificacion-solo-pendientes-plan.md`  
**Registro**: Fila `R-077` en `.agents/plans/REGISTRY.md` (Ronda 376–383, Fases 1788–1797)  
**Estado actual**: `approved` (listo para ejecución con `swarm-parallel-task`)  
**Base commit**: `3c2d1b4` (HEAD actual)  
**Workspace**: `C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`

---

## 1. Contexto y Diagnóstico

En la versión actual de `ConversorPDF.XLL.Academic`, la reclasificación de patrones para redistribución semanal (`FormAutorizarReclasificacion` / `AgEscaladorRanuraAula`) presentaba los siguientes desacoples críticos:
1. `AgFormGeneradorHorarios` obtenía candidatos directamente desde `IdentificarPatronesRedistribuibles(contexto)` **antes** de evaluar el plan fechado, sin consultar si las materias realmente tenían déficit o sesiones pendientes en el plazo.
2. `IdentificarPatronesRedistribuibles` no verificaba el estado del aula (`a.Estado == AgEstadoAula.Disponible`), permitiendo que aulas con estado no disponible (como el caso del aula 12 con `Estado = 4 (Aprobado)`) ingresaran a la lista de candidatas.
3. Al seleccionar cualquier asignatura de dicha aula o con estado desactualizado, `EvaluarInterno` lanzaba un `ArgumentException("Reclasificación ajena o configuración modificada...")`, que se propagaba como `ExcepcionNoControlada` en la interfaz gráfica.
4. Materias que ya caben completamente dentro del plazo figuraban como candidatas a liberar capacidad innecesariamente.

---

## 2. Decisiones de Arquitectura e Invariantes

- **Elegibilidad canónica**: `EsAulaElegible(aula)` exige estrictamente `Id > 0`, `AulaRef` válida y `Estado == AgEstadoAula.Disponible` (1). Aulas en estado Aprobado (4), Inactiva, Mantenimiento o Reservada quedan totalmente excluidas de generar candidatas y de autorizaciones.
- **Candidatas fechadas y acotadas al déficit**: Una asignatura es candidata si y solo si:
  - Pertenece a un aula elegible (`Disponible`).
  - Su origen de patrón es `LEGADO` con días asignados.
  - Presenta sesiones pendientes positivas (`DemandaRequerida > SesionesUbicadas`) en la evaluación fechada del plazo aprobado.
  - Si el aula cabe completamente en plazo (`CabeEnPlazo == true`), devuelve 0 candidatas.
  - Materias completas NUNCA son candidatas a reclasificación, aun si su aula tiene otras asignaturas pendientes.
- **Manejo tipado de rechazos**: Reemplazar excepciones no controladas por validaciones tipadas y mensajes explicativos (WARN en log con IDs y causas; aviso comprensible al usuario en UI).
- **Aislamiento SQL y Enriquecimiento de Decision**: `Decision` expone `IdAula` y en `AplicarDetalladoAsync` la actualización se acota estrictamente con `WHERE id_aula = @IdAula`.
- **Sobrecargas compatibles**: Se mantiene la sobrecarga `IdentificarPatronesRedistribuibles(contexto, aula, evaluacion = null)` para que cada etapa intermedia compile limpiamente.
- **Flujo UI ordenado**:
  1. Capturar contexto y plazo.
  2. Evaluar plan fechado.
  3. Extraer candidatas con déficit.
  4. Si no hay candidatas, omitir `FormAutorizarReclasificacion`.
  5. Si hay candidatas, permitir selección individual desmarcada por defecto.
  6. Simular en copia y avanzar a escalón solo si persiste déficit.
- **Zero Tolerance L0 / Push**: No push a remotos. No modificaciones en archivos L0. Concurrencia máxima: 2 writers con leases disjuntos.

---

## 3. Estructura de Olas y Asignación de Fases (8 Olas / 10 Tareas)

| Wave | Ronda | Tasks | Fases | Enfoque |
|---|---|---|---|---|
| 1 | 376 | T1, T2 | 1788, 1789 | Métricas de modelo y desglose por asignatura con tests en evaluador |
| 2 | 377 | T3, T4 | 1790, 1791 | Calibración preventiva de fixtures de tests y diálogo UI con déficit |
| 3 | 378 | T5 | 1792 | Elegibilidad canónica, candidatas fechadas, IdAula en Decision y aislamiento SQL |
| 4 | 379 | T6 | 1793 | Validación tipada de selección y manejo controlado de rechazos |
| 5 | 380 | T7 | 1794 | Pruebas unitarias de rechazos tipados y casos de aislamiento |
| 6 | 381 | T8 | 1795 | Inversión y orquestación del flujo UI en Generador de Horarios |
| 7 | 382 | T9 | 1796 | Pruebas de aceptación E2E (aula 12, materias completas omitidas, etc.) |
| 8 | 383 | T10 | 1797 | Certificación Gate Full de la suite completa (>= 1259) |

---

## 4. Instrucciones para el Ejecutor

Ejecutar mediante el comando oficial:
```powershell
/swarm-parallel-task .agents/plans/2026-11-14-reclasificacion-solo-pendientes-plan.md
```
El orquestador ejecutará las 8 olas respetando los leases disjuntos, registrando la evidencia de build y pruebas en `_evidence/` y generando commits locales atómicos por tarea.
