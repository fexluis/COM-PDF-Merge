# Handoff: Correcciones Libro 8 / Corrida 94 — Aulas Aprobadas, Autoridad de Conflictos Fechados y Consistencia Staging/UI

> **Fecha**: 2026-11-15  
> **Plan**: `.agents/plans/2026-11-15-libro8-corrida94-plan.md`  
> **Registro**: `R-079` (`REGISTRY.md`) · `Wave 1..6` → `Ronda 392–397` · `T1..T9` → `Fase 1806–1814`

---

## 1. Entorno de Ejecución
- **Workspace**: `C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`
- **Rama**: `codex/huellas-escalado-20261114`
- **Preflight Command**:
  ```powershell
  powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace "C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan ".agents/plans/2026-11-15-libro8-corrida94-plan.md" -AllowOneDrive
  ```
- **Preflight Status**: Aprobado (504 archivos `.cs` en XLL, 182 en Tests, slice Academic verificado, baseline con TOTAL=1387).

---

## 2. Comando de Ejecución
Para iniciar la ejecución multiagente en paralelo:
```text
/swarm-parallel-task .agents/plans/2026-11-15-libro8-corrida94-plan.md
```

---

## 3. Compuerta de Entrada (Gate Obligatorio)
Antes de ejecutar la Wave 1:
1. Invocar al agente `plan-reviewer` sobre `.agents/plans/2026-11-15-libro8-corrida94-plan.md`.
2. Si el dictamen es `BLOCKED` o reporta hallazgos `Critical`, detener inmediatamente la ejecución y subsanar el plan.

---

## 4. Criterios de Aceptación Globales
1. **A1 (Aulas aprobadas protegidas)**: `AgGeneradorHorarios` bloquea la ejecución pre-vuelo con `I5_AULA_APROBADA_CON_DEFICIT` si un aula en estado `Aprobado` no cabe en plazo bajo la ventana efectiva. El escalador no muta aulas aprobadas automáticamente.
2. **A2 (Autoridad común de ocupación fechada y calibración de fixtures)**: `AgCalendarioConflictos` proyecta sesiones con `AgExpansorSesiones` y valida cruces reales por fecha e intervalo horario. `AgValidadorReglas` reemplaza los solapamientos de rangos aproximados por cruces reales (R1, R2, R3, R4, R7), eliminando falsos positivos (e.g., materias intercaladas como Estadística y Teorías en CALI III-A con 0 cruces reales). Se calibran los fixtures existentes en `AgValidadorReglasTests.cs` para registrar bloques válidos en `contexto.BloquesPorId`.
3. **A3 (Disponibilidad grupal docente encapsulada)**: `AgAsignadorTercero` valida disponibilidad (`DisponibleGrupo`) con null-safety sobre todas las fechas del grupo completo contra el individuo y ocupación externa (`contexto.TerceroOcupado`). Encapsula la sincronización del grupo hermano dentro del archivo `AgAsignadorTercero.cs`.
4. **A4 (Persistencia de vacantes y compuerta estricta)**: `AgScheduleRepository.UpdateCorridaAsync` persiste la columna `vacantes`. En la UI (`AgFormGeneradorHorarios`), la corrida solo se considera `factible` y se habilita la promoción si `stagingOk`, cobertura completa y `HardViolations == 0 && Vacantes == 0 && sesionesVacantes == 0`, tanto al finalizar la optimización como al recargar la corrida desde el combo (`CargarCorridaDesdeComboAsync`).
5. **A5 (Métricas no sesgadas)**: `CargarModuloResultadosAsync` muestra conteos de registros y puntuaciones de penalización en campos separados, sin sumar penalizaciones de fitness a recuentos de filas.
6. **A6 (Visibilidad de violaciones)**: `FormVerVinculaciones` muestra violaciones globales bajo nodo raíz dedicado aunque `id_gen` sea nulo y usa cláusula tolerante de subunidad. `SaveIndividuoAsync` resuelve `indicesGenes` a `id_gen` definitivo dentro de la transacción de guardado.
7. **Convenciones de Test (`test-convention.md`)**: Tests de dominio/aplicación usan xUnit 2.7. Tests de formularios bajo `ConversorPDF.Tests/Academic/UI/Forms/` (`FormVerVinculacionesViolacionesTests`, `AgFormGeneradorHorariosResultadosTests`) usan MSTest 3.3.1 en hilo STA (`[STATestMethod]`).
8. **Regresiones**: La suite completa pasa en verde con `scripts/verify.ps1 -Mode Full` (0 fallos, 0 omitidos, `TOTAL >= 1391`).

---

## 5. Prohibiciones y Restricciones
- ❌ **No push / no merge**: No intentar enviar cambios a repositorios remotos (`no_push: true`).
- ❌ **No C# 12+ syntax**: El código debe compilar en .NET Framework 4.8 / C# 7.3 (sin expresiones de colección `[]`, sin constructores primarios para clases de DI).
- ❌ **No tocar hot paths L0 sin rol `integrator`**: `ConversorPDF.XLL.csproj` se modifica únicamente en T1 por `integrator`.
- ❌ **No saltar tests**: Cada tarea debe cumplir su comando de verificación antes de pasar a la siguiente ola.

---

## 6. Entregables Esperados
- Registro de tareas y olas actualizado en el ledger del plan.
- Commits locales atómicos por cada tarea completada.
- Informes de verificación generados bajo `_evidence/<timestamp>/summary.md`.
- Cierre y sellado del plan en `REGISTRY.md` mediante `swarm-verify`.
