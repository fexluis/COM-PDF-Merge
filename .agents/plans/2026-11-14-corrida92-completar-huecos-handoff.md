# Handoff: Corrida 92 — Completar Huecos y Unificar Evaluación y Generación

1. **Workspace a abrir**:
   - `C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72` (con bandera `-AllowOneDrive`).

2. **Rama y estado pre-flight**:
   - Rama: `codex/huellas-escalado-20261114`
   - Registro: Fila `R-078` en `.agents/plans/REGISTRY.md` (Rondas 384–391, Fases 1798–1805).
   - Baseline: `tests/baseline/test_baseline.txt` (TOTAL: 1359).
   - Pre-flight: `powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -AllowOneDrive -Workspace "C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan ".agents/plans/2026-11-14-corrida92-completar-huecos-plan.md"`.

3. **Comando de ejecución**:
   - `/swarm-parallel-task .agents/plans/2026-11-14-corrida92-completar-huecos-plan.md` (o fallback secuencial con `executing-plans-local`).

4. **Gate obligatorio**:
   - Invocación de `plan-reviewer` previo al inicio de Wave 1. Plan revisado y blindado contra deadlocks de test y solapes de lease.

5. **Criterios de aceptación globales**:
   - Conservar las restricciones no reclasificadas (patrones LEGADO/FIJO) y completar materias automáticas con sesiones fechadas libres aprovechando huecos entre bandas.
   - Paridad estricta entre `Construir` y `ConstruirEvaluacion`: mismas sesiones y fechas ante idénticas entradas.
   - Desacople completo de constantes rígidas de fechas (eliminación de `limiteMayo = 30/05/2027` sin detener compactación).
   - Conservación de `FechasGen` exactas en la fábrica y en staging sin desajustes por redistribución en `AgCoherenciaGrupo.Sellar`.
   - Transiciones RED -> GREEN comprobadas por tarea con suite de regresión aislada (`AgPlanificadorRanurasRegresionTests.cs`).
   - 0 regresiones contra el baseline de 1359 tests existentes.

6. **Prohibiciones**:
   - ❌ No realizar push ni merge a remoto (`no_push: true`).
   - ❌ No usar sintaxis C# 12+ (colecciones `[]`, primary constructors para clases de DI).
   - ❌ No modificar archivos de hot paths L0 (`ConversorPDF.XLL.csproj`, `App.cs`, migraciones, etc.).
   - ❌ No omitir ni debilitar tests unitarios.

7. **Entregables esperados**:
   - Plan `.agents/plans/2026-11-14-corrida92-completar-huecos-plan.md` con ledger completo en estado `done`.
   - Commits locales atómicos por ola/tarea.
   - Directorios de evidencia en `_evidence/<timestamp>/summary.md`.
