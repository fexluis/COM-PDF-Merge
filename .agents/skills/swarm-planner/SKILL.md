---
name: swarm-planner
description: >
  [EXPLICIT INVOCATION ONLY] Creates dependency-aware implementation plans optimized for parallel
  multi-agent execution.
metadata:
  invocation: explicit-only
---

# Swarm-Ready Planner

Crea planes de implementación con dependencias explícitas, optimizados para ejecución multiagente en paralelo. Esta skill puede ejecutarse dentro o fuera de Plan Mode y **nunca implementa código**: solo planifica.

## 1. Propósito y autoridad del contrato

- **Salida obligatoria**: `.agents/plans/<YYYY-MM-DD>-<slug>-plan.md` + `.agents/plans/<YYYY-MM-DD>-<slug>-handoff.md`.
- **Numeración global**: [`plans/REGISTRY.md`](../../plans/REGISTRY.md) — reserva de `Ronda` (=`Wave`) y `Fase` (=`T<n>`). Sin fila en el registro, el plan no es ejecutable por la Fase C.
- **Fuente única de verdad**: [`references/plan-contract.md`](references/plan-contract.md). Ante conflicto con este documento, **manda el contrato**.
- **Reglas siempre activas**: [`.agents/rules/parallel-work.md`](../../rules/parallel-work.md) — leases L0/L1/L2, `max_parallel_writers: 2` y gate serializado. No son negociables.
- **Consumidores del plan**: `swarm-parallel-task` (orquestador con subagentes) o `executing-plans-local` (fallback secuencial cuando no hay `invoke_subagent`).

## 2. Core Principles

1. **Explore Codebase**: investiga arquitectura, patrones, implementaciones existentes, dependencias y frameworks en uso.
2. **Fresh Documentation First**: usa Context7 para CUALQUIER librería, framework o API externa antes de planificar tareas.
3. **Ask Questions**: aclara ambigüedades de alcance, restricciones o prioridades en cualquier momento del proceso.
4. **Explicit Dependencies**: cada tarea declara de qué depende, habilitando el máximo paralelismo.
5. **Atomic Tasks**: cada tarea es ejecutable de forma independiente por un único agente.
6. **Disjoint Leases**: dos tareas de la misma ola nunca comparten archivos ni prefijos en `files_lock`.
7. **Review Before Yield**: autoevaluación de 13 puntos (§3.8) y revisión con `plan-reviewer` (§3.9) antes de finalizar.

## 3. Process

### 3.1 Research

- Arquitectura, patrones e implementaciones existentes relevantes al slice objetivo.
- Dependencias y frameworks realmente presentes en el repositorio (no supuestos).
- Si existe un context pack para el slice, léelo antes de abrir código: presupuesto ≤ 400 líneas (`scripts/context-pack.ps1 -Slice <X>`).

### 3.2 Optional: Stop to Clarification Questions

- Si la arquitectura es desconocida o falta información, **DETÉN y DEVUELVE EL TURNO** al usuario con preguntas antes de continuar.
- Ofrece siempre una recomendación para cada pregunta.
- Si la arquitectura está clara, continúa en §3.3.

### 3.3 Documentation

**Obligatorio para dependencias externas.** Usa la skill o MCP Context7 para obtener documentación vigente de toda librería, framework o API que se use o vaya a usarse. Si Context7 no está disponible, usa búsqueda web.

Objetivo: APIs y parámetros correctos para la versión real del proyecto.

### 3.4 STOP and Request User Input

Ante cualquier decisión que admita más de una vía razonable:

- Detente y pregunta de inmediato.
- No asumas alcance, restricciones ni prioridades.
- Las preguntas deben eliminar ambigüedad y reducir riesgo.
- Ofrece siempre una recomendación.
- Usa la herramienta de preguntas del entorno (`ask_question` / `request_user_input` / `AskUserQuestion`, según disponibilidad).

### 3.5 Create Dependency-Aware Plan

Cada tarea DEBE declarar los 11 campos de `references/plan-contract.md` §3, en este orden:

| Campo | Contenido |
|---|---|
| `id` | Identificador único (`T1`, `T2`) |
| `depends_on` | IDs que deben completarse antes (`[]` si es raíz) |
| `location` | Carpeta o archivo principal de trabajo |
| `files_lock` | Lista exacta de archivos; disjunta dentro de la misma ola |
| `description` | Qué logra la tarea, acotado |
| `validation` | Criterio falsable y observable |
| `verification` | Comando exacto: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "..."` |
| `evidence` | Ruta de evidencia que genera `build-verifier`: `_evidence/<timestamp>/summary.md` |
| `status` | `pending`, `executing`, `red`, `green`, `blocked`, `done` (ningún otro valor) |
| `log` | Registro de ejecución que escribe el orquestador |
| `files edited/created` | Archivos finales modificados, verificados contra `files_lock` |

**Campo adicional `fase` (numeración global)**: obligatorio si el plan tiene fila en `REGISTRY.md`. Se declara como `- **fase**: <n>` justo después de `- **id**:` y equivale a la `Fase {n}` de `.clinerules/workflows/Ronda.md`. La tabla de olas lleva además una columna `Ronda` (§3.10).

**Caminos calientes (Lease L0)** — asignables SOLO al rol `integrator`, serializado y una tarea a la vez:

- `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- `ConversorPDF.XLL/Core/Helpers/App.cs`
- `ConversorPDF.XLL/**/Migrations/MigrationV*.cs`
- `ConversorPDF.XLL/**/Persistence/SchemaMigrator.cs`
- `ConversorPDF.XLL/Ribbon/RibbonController.cs`
- Composition Root / contenedor IoC-DI
- `docs/toc.yml`

**Tipos nuevos ⇒ `integrator-first`**: si una tarea introduce clases o interfaces nuevas, emite antes una tarea de `integrator` que cree el esqueleto con `NotImplementedException` y lo registre en el `.csproj`. En C# el test RED debe compilar.

**Regla de olas**: si dos tareas comparten archivo o prefijo en `files_lock`, van en olas distintas o se marcan como serializadas. Tope global `max_parallel_writers: 2`.

**Ruling (solo en disputa)**: `task-reviewer` puede rechazar hasta 3 rondas de fix; si persiste el desacuerdo, registra `Ruling: <decisión> — <justificación técnica> — <costo si la decisión es errónea>`.

### 3.6 Save Plan

> [!CRITICAL]
> Guardar OBLIGATORIAMENTE en `.agents/plans/<YYYY-MM-DD>-<slug>-plan.md`. NUNCA en la raíz del repositorio.

### 3.7 Write Handoff

> [!CRITICAL]
> Escribe también `.agents/plans/<YYYY-MM-DD>-<slug>-handoff.md` (plantilla en [`.agents/plans/README.md`](../../plans/README.md) §4). El subagente **no hereda contexto**: el handoff es su único punto de entrada.

Contenido mínimo del handoff:

1. Workspace a abrir (copia del árbol real, fuera de OneDrive).
2. Rama y estado del pre-flight (`git status --porcelain` vacío).
3. Comando de ejecución: `/swarm-parallel-task .agents/plans/<plan>.md`.
4. Gate obligatorio: invocación inicial de `plan-reviewer`; `BLOCKED` o hallazgo `Critical` ⇒ abortar.
5. Criterios de aceptación globales.
6. Prohibiciones (push/merge, C# 12, tocar L0 sin `integrator`, saltar tests).
7. Entregables esperados (ledger completo, commits locales, `_evidence/`).

### 3.8 Self-Review: checklist de 13 puntos

Antes de ceder el turno, evalúa el plan contra estos 13 puntos y **corrige lo que falle**:

| # | Punto |
|---|---|
| 1 | Dependencias completas: ningún `depends_on` omitido |
| 2 | Orden ejecutable: cada tarea compila al terminar |
| 3 | Ninguna tarea mezcla L0 sin rol `integrator` |
| 4 | `files_lock` disjuntos por ola |
| 5 | Cada `verification` es un comando que existe en el repo |
| 6 | Cada `validation` es falsable |
| 7 | El RED compila (esqueleto `integrator-first` previsto) |
| 8 | Migración y `APP_VERSION` asignadas solo por `integrator` |
| 9 | Ninguna tarea cruza más de un slice |
| 10 | El context pack usado tiene ≤ 400 líneas |
| 11 | Criterio de aceptación global definido |
| 12 | Riesgos con mitigación |
| 13 | `numbering:` en el envelope y fila `reserved` añadida a `REGISTRY.md` |

Registra el resultado dentro del plan (tabla idéntica + nota de correcciones aplicadas).

### 3.9 Subagent Review

Tras guardar plan y handoff, revisa con el rol `plan-reviewer` (`.agents/agents/plan-reviewer.md`):

- Si el entorno dispone de `invoke_subagent`, invócalo con esta consigna.
- Si no existe, aplica tú mismo las 9 reglas duras de `plan-reviewer.md` y emite el veredicto.

```text
Revisa este plan de implementación:
1. Dependencias faltantes entre tareas
2. Problemas de orden que causarían fallos
3. Falta de manejo de errores o casos borde
4. Huecos, agujeros y gotchas

Entrega feedback accionable. No hagas preguntas.
Veredicto: APROBADO | BLOCKED

Ubicación del plan: [ruta]
Contexto: [breve]
```

Si el veredicto es `BLOCKED` con retroalimentación accionable, revisa el plan **antes** de ceder. Un plan `BLOCKED` no se ejecuta.

### 3.10 Reserva de numeración (Ronda/Wave · Fase/Task)

`.clinerules/workflows/Ronda.md` numera las revisiones como `Ronda {n}` / `Fase {n}`; los planes swarm usan `Wave {n}` / `T{n}`. Es el **mismo objeto** con dos vocabularios, y la numeración es **global y monótona**: vive en [`plans/REGISTRY.md`](../../plans/REGISTRY.md) (contrato en [`plan-contract.md`](references/plan-contract.md) §7).

Procedimiento obligatorio al crear el plan:

1. **Leer contadores** (§2 del registro):

   ```powershell
   $reg   = Get-Content .agents\plans\REGISTRY.md -Raw
   $ronda = [int]([regex]::Match($reg, '\|\s*`ronda`\s*\|\s*(\d+)\s*\|').Groups[1].Value)
   $fase  = [int]([regex]::Match($reg, '\|\s*`fase`\s*\|\s*(\d+)\s*\|').Groups[1].Value)
   "$ronda / $fase"   # siguiente = $ronda+1 / $fase+1
   ```

2. **Calcular la reserva**: `W` olas ⇒ `ronda_base = ronda + 1`, rango `[ronda_base, ronda_base + W - 1]`; `N` tareas ⇒ `fase_base = fase + 1`, rango `[fase_base, fase_base + N - 1]`.
3. **Declarar en el envelope** el bloque `numbering:` y, en cada tarea, `- **fase**: <n>` (§4).
4. **Añadir la fila `reserved`** en §3 del registro: `R-00x`, fecha, archivo del plan, `Wave 1..W → Ronda a–b`, `T1..TN → Fase c–d`, `verify: pendiente`, `HEAD` actual.
5. **No tocar §2** (contadores): se actualizan al cerrar, en la Fase C (`swarm-verify`), con el último número **usado**.
6. **Re-planificación**: conserva los números ya reservados. Solo se amplía el rango si se **añaden** tareas al final; nunca se compacta ni se reordena.
7. **Plan de fix** (`<fecha>-<slug>-fix-plan.md`): rango **nuevo** + referencia explícita a la fila y a las `Fase {n}` que corrige.
8. **Olas reales ≠ previstas**: la Fase B puede serializar más olas por colisión de leases; el cierre ajusta el rango al alza y anota la desviación (`LEASE-SERIALIZED`).

Un plan sin bloque `numbering:` ni fila en el registro es **`BLOCKED`** en la Fase C: no hay forma de sellar la numeración.

## 4. Plan Template

Bloque literal alineado 1:1 con `references/plan-contract.md`. Copiar y rellenar; no inventar campos ni renombrar encabezados.

```markdown
# Plan: <Título descriptivo del feature o refactor>
**Generado**: <YYYY-MM-DD> · **Autor**: Cline/Swarm-Planner · **Estado**: draft | approved | executing | closed

## Execution envelope
- branch: refactor/<slug>
- workspace: <copia del árbol real VALIDADA, fuera de OneDrive>
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <workspace> -Plan <plan>
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: <Slice>
- baseline: tests/baseline/test_baseline.txt (TOTAL: <medido en el workspace, NUNCA estimado>)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0:
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Core/Helpers/App.cs
  - ConversorPDF.XLL/**/Migrations/MigrationV*.cs
  - ConversorPDF.XLL/**/Persistence/SchemaMigrator.cs
  - ConversorPDF.XLL/Ribbon/RibbonController.cs
  - Composition Root / IoC Container
  - docs/toc.yml
- stop_conditions:
  - Acción destructiva no reversible sin respaldo
  - Violación de seguridad o datos
  - Intento de push, merge o publicación a remoto
  - Plan en estado irrecuperable (más de 3 rondas de fix fallidas)
- ledger_owner: swarm-parallel-task
- no_push: true
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: <R-00x>
  - ronda_base: <n>          # Wave 1 = Ronda <n>
  - fase_base: <n>           # T1 = Fase <n>
  - waves: <W>
  - tasks: <N>

## Dependency graph
T1 ── T3 ───┐
      │     └── T5 ── T6
T2 ── T4 ───┘

## Tasks

### T1: <Nombre>
- **id**: T1
- **fase**: <fase_base>
- **depends_on**: []
- **location**: <ruta o carpeta>
- **files_lock**: [<archivos exactos>]
- **description**: <qué hace>
- **validation**: <criterio falsable>
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "..."
- **evidence**: (la llena build-verifier)
- **status**: pending
- **log**: (lo llena el orquestador)
- **files edited/created**: (lo llena el orquestador)

### T2: <Nombre>
- **id**: T2
- **depends_on**: []
- **location**: <ruta o carpeta>
- **files_lock**: [<archivos exactos>]
- **description**: <qué hace>
- **validation**: <criterio falsable>
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "..."
- **evidence**: (la llena build-verifier)
- **status**: pending
- **log**: (lo llena el orquestador)
- **files edited/created**: (lo llena el orquestador)

### T3: <Nombre>
- **id**: T3
- **depends_on**: [T1]
- **location**: <ruta o carpeta>
- **files_lock**: [<archivos exactos>]
- **description**: <qué hace>
- **validation**: <criterio falsable>
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "..."
- **evidence**: (la llena build-verifier)
- **status**: pending
- **log**: (lo llena el orquestador)
- **files edited/created**: (lo llena el orquestador)

## Parallel execution groups
| Wave | Ronda | Tasks | Can start when |
|---|---|---|---|
| 1 | <ronda_base> | T1, T2 | Inmediato (leases disjuntos) |
| 2 | <ronda_base + 1> | T3 | T1 y T2 completados |

## Testing strategy
- <qué se prueba y con qué comando>

## Risks & mitigations
- <qué puede fallar y cómo se mitiga>
```

## 5. Prohibiciones y recordatorios

- **No implementar**: esta skill solo crea el plan; el código lo escriben `swarm-parallel-task` o `executing-plans-local`.
- **No escribir el plan en la raíz** del repositorio: solo `.agents/plans/`.
- **No reutilizar números** de `Ronda`/`Fase`: la reserva se hace en `REGISTRY.md` y es append-only; nunca se bajan los contadores ni se compactan rangos.
- **No asignar Lease L0** a `implementer` ni a `tdd-red-writer`.
- **No inventar el `TOTAL` del baseline**: se mide con `scripts/verify.ps1 -Mode Full`; un `TOTAL` menor que el baseline es `SUITE_SHRUNK` (exit 3), no un éxito.
- **No planificar ejecución sobre OneDrive**: `scripts/preflight.ps1` devuelve 3; usar `-AllowOneDrive` solo con aviso explícito en el ledger.
- **No push, merge ni publish**: el swarm hace commits locales atómicos y `no_push: true`.
- **Siempre** Context7 para dependencias externas antes de cerrar el plan.
- **Siempre** preguntar donde haya ambigüedad; nunca asumir alcance.

