# Directorio de Planes Swarm (`.agents/plans/`)

Este directorio contiene los planes de ejecución para características, refactorizaciones y tareas atómicas coordinadas entre Cline y los subagentes de Antigravity.

---

## 1. Reglas de Nombrado y Estructura
- Formato de archivo de plan: `<YYYY-MM-DD>-<slug>-plan.md`
- Formato de handoff complementario: `<YYYY-MM-DD>-<slug>-handoff.md`
- **Prohibición**: Ningún plan debe crearse en la raíz del repositorio.

---

## 2. Ownership y Concurrencia de Escritura (Regla de un solo escritor)
- **Fase A (Planificación)**: Cline es el único escritor. Produce el plan y el handoff.
- **Fase B (Ejecución)**: El orquestador de Antigravity (`swarm-parallel-task`) es el único escritor. Actualiza el ledger y los estados `T<n>`. Cline trata el plan como solo lectura.
- **Fase C (Cierre)**: Cline re-ejecuta verificación, valida invariantes y sella el plan como `closed`. Si se requieren correcciones, se crea un plan de fix independiente.

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
