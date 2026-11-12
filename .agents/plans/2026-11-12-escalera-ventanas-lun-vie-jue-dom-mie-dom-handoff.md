# Handoff: Escalado Automático de Ventanas LUN-VIE / JUE-DOM / MIE-DOM (R-058)

**Fecha**: 2026-11-12  
**Plan asociado**: `.agents/plans/2026-11-12-escalera-ventanas-lun-vie-jue-dom-mie-dom-plan.md`  
**Registro**: Fila `R-058` en `.agents/plans/REGISTRY.md` (Ronda 345–350, Fases 1707–1714)  
**Estado actual**: `closed` (implementación terminada; revisión Ronda 351 detectó y aplicó correcciones pendientes)  
**Commit implementado**: `1a97cfc`  
**Commit verificado**: `27b7c71`; Debug/x64; Build/Fast/Full (1236 PASS)  
**Release y smoke**: pendientes

---

## 1. Contexto y Problema

El planificador académico y el escalador de ventanas (`AgEscaladorVentanaAula`) operaban con una escalera restringida (`LUN-JUE`, `MAR-VIE`, `VIE-SAB`, `JUE-SAB`, `"4,5,6,7"`).
Se incorporaron de forma canónica las ventanas extendidas:
- `LUN-VIE` ("1,2,3,4,5") como escalón de `LUN-JUE` y `MAR-VIE`.
- `JUE-DOM` ("4,5,6,7") y `MIE-DOM` ("3,4,5,6,7") como escalones sucesivos de `VIE-DOM`.

---

## 2. Resumen de Olas y Cierre

- **Wave 1..6 (Ronda 345–350 / Fases 1707–1714)**:
  - T1..T8: Completadas y selladas con estado `done`.
  - T2 y T3 conservan su evidencia RED histórica certificada en `_evidence/`.
  - Gate final `-Mode Full` certificado con 1223 tests PASS (baseline previo: 1172).

---

## 3. Estado de Revisión Ronda 351 (Fases 1715–1719)

- **Fase 1715**: Verificación de filas persistidas (`aplicadas != esperadas`) antes de mutar estado en memoria en `AgFormGeneradorHorarios.cs`.
- **Fase 1716**: Exclusión mutua con `SemaphoreSlim` entre sincronización debounce y escalado; bloqueo de árbol y combo con `try-finally`; `RecargarDimensionesPerfilActivoTrasEscaladoAsync` devuelve `Task<bool>` y aborta ante inconsistencia.
- **Fase 1717**: Lectura de persistencia real en `AgEscaladorVentanaAulaPersistenciaTests.cs` y adición de test de integración con repositorio/mapper.
- **Fase 1718**: Rechazo estricto de tokens de franja `6M/6T` en listas en `AgDiasVentana.cs` y tests en `AgDiasVentanaTests.cs`.
- **Fase 1719**: Sincronización documental en handoff, plan y registro.

---

## 4. Re-gate post-cierre (Ronda 357 / Fase 1733)

- **Build**: PASS. Se añadió el `using ConversorPDF.XLL.Academic.Application.DTOs` faltante en `AgEscaladorVentanaAulaPersistenciaTests.cs` (la corrección R-061 quedaba incompleta y rompía el build del proyecto de tests).
- **Fast de aceptación**: 125/125 PASS (`AgEscaleraVentanas`, `AgDiasVentana`, `AgEscaladorVentanaAula`, `AgPlanificadorOleadas`, `AgNodoArbolKeyMapper`).
- **Full**: 1236/1236 PASS, 0 fallos, exit 0 (baseline 1223). Evidencia: `_evidence/20261112-170715`.
- **Correcciones R-059..R-064**: verificadas por el gate; commit `27b7c71`.
- **Release y smoke**: siguen pendientes (fuera del alcance aprobado).

