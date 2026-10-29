# Directorio de Planes Swarm (`.agents/plans/`)

Este directorio contiene los planes de ejecución para características, refactorizaciones y tareas atómicas coordinadas entre Cline y los subagentes de Antigravity.

---

## 1. Reglas de Nombrado y Estructura
- Formato de archivo de plan: `<YYYY-MM-DD>-<slug>-plan.md`
- Formato de handoff complementario: `<YYYY-MM-DD>-<slug>-handoff.md`
- **Registro de numeración**: [`REGISTRY.md`](REGISTRY.md) — contadores y asignaciones de `Ronda` (=`Wave`) y `Fase` (=`T<n>`). Es la fuente única y es **append-only**.
- **Prohibición**: Ningún plan debe crearse en la raíz del repositorio.

---

## 2. Ownership y Concurrencia de Escritura (Regla de un solo escritor)
- **Fase A (Planificación)**: Cline es el único escritor. Produce el plan y el handoff.
- **Fase B (Ejecución)**: El orquestador de Antigravity (`swarm-parallel-task`) es el único escritor. Actualiza el ledger y los estados `T<n>`. Cline trata el plan como solo lectura.
- **Fase C (Cierre)**: Cline re-ejecuta verificación, audita el código de cada `T<n>` (`scripts/audit-tasks.ps1`), valida invariantes, cierra la fila del plan en [`REGISTRY.md`](REGISTRY.md) y sella el plan como `closed`. Si se requieren correcciones, se crea un plan de fix independiente.

---

## 3. Plantilla Oficial de Plan (`plan-template.md`)

> Fuente única del contrato: [`swarm-planner/references/plan-contract.md`](../skills/swarm-planner/references/plan-contract.md).
> Esta plantilla es un resumen; ante conflicto, manda `plan-contract.md`.

```markdown
# Plan: <Título>
**Generado**: <YYYY-MM-DD> · **Autor**: Cline/Swarm-Planner · **Estado**: draft|approved|executing|closed

## Execution envelope
- branch: refactor/<slug>
- workspace: <copia del arbol real, fuera de OneDrive>
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ruta> -Plan <plan>
- workspace_guard: { min_xll_cs: 300, min_tests_cs: 100, required_slice: Academic }
- baseline: tests/baseline/test_baseline.txt (TOTAL: <medido, nunca estimado>)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0: [ConversorPDF.XLL.csproj, Core/Helpers/App.cs, **/Migrations/**, SchemaMigrator.cs, RibbonController.cs, composition root, docs/toc.yml]
- stop_conditions: [destructivo, seguridad, push/merge/publish, plan irrecuperable tras 3 fixes]
- ledger_owner: swarm-parallel-task
- no_push: true

## Dependency graph
```
[Grafo ASCII o Mermaid]
```

## Tasks
### T1: <Nombre>
- **depends_on**: []
- **location**: <Ruta>
- **files_lock**: [<Archivos exactos>]
- **description**: <Qué hace>
- **validation**: <Criterio falsable>
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "..."
- **evidence**: _evidence/<timestamp>/summary.md
- **status**: pending
- **log**:
- **files edited/created**:

## Parallel execution groups
| Wave | Tasks | Can start when |
|---|---|---|
| 1 | T1 | Immediately |

## Testing strategy
## Risks & mitigations
```

---

## 4. Plantilla Oficial de Handoff (`handoff-template.md`)

```markdown
# Handoff: <Título de la feature>
1. **Workspace a abrir**: `C:\dev\ConversorPDF` (nunca OneDrive).
2. **Rama y estado pre-flight**: `git status --porcelain` limpio en `refactor/<slug>`.
3. **Comando de ejecución**: `/swarm-parallel-task .agents/plans/<YYYY-MM-DD>-<slug>-plan.md`.
4. **Gate obligatorio**: Invocación inicial de `plan-reviewer`. Si hay hallazgos Critical, abortar.
5. **Criterios de aceptación globales**: RED->GREEN por tarea, cero regresiones vs baseline, cero commits fuera de lease.
6. **Prohibiciones**: push/merge, C# 12, tocar L0 sin integrator, saltar tests.
7. **Entregables esperados**: Plan con ledger completo + commits locales + directorio `_evidence/`.
```

---

## 5. Plantilla Oficial de Cierre (Fase C)

La Fase C la ejecuta Cline y **no acepta el ledger del orquestador tal cual**: re-verifica en disco, valida invariantes y solo entonces sella el plan como `closed`.

### 5.1 Runbook (C0 → C6)

```powershell
git status --porcelain                                                   # C0: debe estar VACIO
Get-Content .agents\plans\<plan>.md | Select-String 'status:'            # C0: todos los T<n> en done
powershell -ExecutionPolicy Bypass -File scripts\verify.ps1 -Mode Full   # C1: exit 0 y TOTAL >= baseline
powershell -ExecutionPolicy Bypass -File scripts\audit-tasks.ps1 -Plan .agents\plans\<plan>.md   # C1.5: exit 0
Get-Content .agents\plans\REGISTRY.md | Select-String '^\| `(ronda|fase)`|R-0'                   # C2.5: reserva vigente
git log --oneline <base>..HEAD                                           # C2: un commit por tarea
git log --oneline origin/refactor/<slug>..HEAD                           # C4: los commits locales NO se publicaron
```

Exit codes de `scripts\audit-tasks.ps1`: `0` sin hallazgos bloqueantes · `3` ≥1 hallazgo `Compilation`/`Runtime`/`Logic` · `4` entrada inválida. Evidencia: `_evidence/<stamp>-audit/audit.md` (con tabla `| Fase | Task | Severidad | Archivo | Líneas |`).

Exit codes de `scripts\verify.ps1`: `0` OK · `1` build falló · `2` tests fallaron / infraestructura · `3` regresión vs baseline, `SUITE_SHRUNK` o `NO_BASELINE` · `4` tooling o workspace inválido.

Invariantes de arquitectura (esperado: 0 resultados):

```powershell
Select-String -Path ConversorPDF.XLL\Academic\Application\**\*.cs -Pattern 'Excel\.|Microsoft\.Office|DataGridView'
Select-String -Path ConversorPDF.XLL\Academic\**\*.cs -Pattern 'new \[\]|DateTime\.Now|\.Result\b|\.Wait\(\)'
```

### 5.2 Tabla de decisión

| verify Full | Leases | Arquitectura | Acción |
|---|---|---|---|
| exit 0 y `TOTAL >= baseline` | OK | OK | **Sellar** `closed` |
| exit 0 | violación de lease | cualquiera | plan `T-fix` |
| exit 0 | OK | hallazgo de arquitectura | plan `T-fix` |
| exit 3 (`SUITE_SHRUNK` / regresión) | cualquiera | cualquiera | plan `T-fix` |
| exit 1, 2 o 4 | — | — | `BLOCKED`: reportar al humano |

### 5.3 Bloque de cierre a añadir al plan

```markdown
## Cierre (Fase C) — YYYY-MM-DD
- **verify Full**: exit 0 · TOTAL <medido> / PASSED <n> / FAILED 0 (baseline <N>)
- **Auditoría de código**: audit-tasks exit 0 · <n> archivos · <n> hallazgos (<w> warnings) · _evidence/<stamp>-audit/audit.md
- **Numeración**: Ronda <r0>–<r1> (Wave 1..<w>) · Fase <f0>–<f1> (T1..<n>) · REGISTRY.md fila <R-xxx>
- **HEAD**: <sha> · **Rama**: refactor/<slug>
- **Evidencia**: _evidence/<timestamp>/summary.md
- **Leases**: 0 archivos fuera de files_lock en <n> commits
- **Arquitectura**: greps de 5.1 sin resultados
- **Publicación**: sin push (origin/<slug> sin cambios)
- **Docs**: docs/articles/changes/YYYY-MM-DD-<slug>.md + walkthrough.md
- **Desviaciones**: <ninguna | Ruling: ...>
```

El único cambio en la cabecera es `**Estado**: closed`. Ciclo de vida: `draft` → `approved` → `executing` (Fase B) → `closed` (Fase C).

### 5.4 Regla de oro (si hay correcciones)

Si la re-verificación encuentra hallazgos, **no se reabre el plan original** (dos escritores sobre el ledger). Se crea `.agents/plans/<YYYY-MM-DD>-<slug>-fix-plan.md` con tareas `T-fix` que referencian el original. Precedente: `2026-10-27-usuarios-duplicados-subunidad-plan.md` §12 y su plan `-fix`.

---

## 6. Numeración y trazabilidad con `Ronda.md`

Los planes swarm y la revisión por rondas (`.clinerules/workflows/Ronda.md`) describen el mismo objeto:

| `Ronda.md` | Plan swarm | Nivel |
|---|---|---|
| `Ronda {n}: {título}` | `Wave {n}` (parallel execution group) | Grupo paralelo |
| `Fase {n}: {título}` | `T{n}: {Nombre}` | Tarea atómica |

`Ronda`/`Fase` son números **globales y monótonos** y viven en [`REGISTRY.md`](REGISTRY.md); `Wave`/`T{n}` son **locales del plan** y se mapean con `numbering:` del envelope (`ronda_base`, `fase_base`) y el campo `- **fase**:` de cada tarea.

| Momento | Quién | Qué escribe |
|---|---|---|
| Fase A | `swarm-planner` | bloque `numbering:` + `fase` por tarea + fila `reserved` |
| Fase B | `swarm-parallel-task` | no renumera; `LEASE-SERIALIZED` si mueve tareas de ola |
| Fase C | `swarm-verify` | valida rangos, fila `closed`, contadores §2 y bloque de cierre |
| Revisión | Cline (`Ronda.md`) | hallazgos como `Fase {n} (T{k})` usando el número global |

Un plan sin fila en el registro es `BLOCKED` en la Fase C; un hallazgo sin `Fase` global se reporta con `T{k}` y abre incidencia de numeración.
