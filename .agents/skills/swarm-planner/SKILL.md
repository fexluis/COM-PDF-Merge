---
name: swarm-planner
description: >
  [EXPLICIT INVOCATION ONLY] Creates dependency-aware implementation plans optimized for parallel
  multi-agent execution.
metadata:
  invocation: explicit-only
---

# Swarm-Ready Planner

Create implementation plans with explicit task dependencies optimized for parallel agent execution. This skill can be ran inside or outside of Plan Mode. 

## Core Principles

1. **Explore Codebase**: Investigate architecture, patterns, existing implementations, dependencies, and frameworks in use.
2. **Fresh Documentation First**: Use Context7 for ANY external library, framework, or API before planning tasks
3. **Ask Questions**: Clarify ambiguities and seek clarification on scope, constraints, or priorities throughout the planning process. At any time.
4. **Explicit Dependencies**: Every task declares what it depends on, enabling maximum parallelization
5. **Atomic Tasks**: Each task is independently executable by a single agent
6. **Review Before Yield**: A subagent reviews the plan for gaps before finalizing

## Process

### 1. Research

**Codebase investigation:**
- Architecture, patterns, existing implementations
- Dependencies and frameworks in use

### 1a. Optional: Stop to Clarification Questions

- If the architecture is unclear or missing STOP AND YIELD to the user, and request user input (AskUserQuestions) before moving on. Always offer recommendations for clarification questions.
- If architecture is present, skip 1a and move onto next step. 

### 2. Documentation

**Documentation retrieval (REQUIRED for external dependencies):**

Use Context7 skill or MCP to fetch current docs for any libraries/frameworks or APIs that are or will be used in project. If Context7 is not available, use web search.

This ensures version-accurate APIs, correct parameters, and current best practices.

### 3. STOP and Request User Input

When anything is unclear or could reasonably be done multiple ways:
- Stop and ask clarifying questions immediately
- Do not make assumptions about scope, constraints, or priorities
- Questions should reduce risk and eliminate ambiguity
- Always offer recommendations for clarification questions.
- Use request_user_input or AskUserQuestion tool if available. 

### 4. Create Dependency-Aware Plan

Structure the plan with explicit task dependencies using this format:

#### Task Dependency Format

Cada tarea DEBE incluir los campos definidos en `references/plan-contract.md`:
- **id**: Identificador único (`T1`, `T2`)
- **depends_on**: Array de IDs previos (`[]` para tareas raíz)
- **location**: Carpeta o archivo
- **files_lock**: Lista exacta de archivos a tocar (deben ser disjuntos por ola)
- **description**: Qué logra la tarea
- **validation**: Criterio de aceptación observable
- **verification**: Comando exacto de verificación (ej. `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "..."`)
- **evidence**: Ruta de evidencia generada (`_evidence/<timestamp>/summary.md`)
- **status**: `Not Completed | In Progress | Completed`
- **log**: Registro de ejecución
- **files edited/created**: Archivos finales modificados

**Importante sobre caminos calientes (Lease L0)**:
Archivos como `ConversorPDF.XLL.csproj`, `Core/Helpers/App.cs`, `**/Migrations/**`, `SchemaMigrator.cs`, `RibbonController.cs` y composition root SOLO pueden ser asignados al rol `integrator` de manera serializada.

### 4. Save Plan

> [!CRITICAL]
> Guardar OBLIGATORIAMENTE en `.agents/plans/<YYYY-MM-DD>-<slug>-plan.md`. NUNCA guardar en la raíz del repositorio.

### 5. Subagent Review

After saving, spawn a subagent to review the plan:

```
Review this implementation plan for:
1. Missing dependencies between tasks
2. Ordering issues that would cause failures
3. Missing error handling or edge cases
4. Gaps, holes, gotchas.

Provide specific, actionable feedback. Do not ask questions.

Plan location: [file path]
Context: [brief context about the task]
```

If the subagent provides actionable feedback, revise the plan before yielding.


## Plan Template

Consulta la especificación completa en [plan-contract.md](file:///c:/Users/USUARIO/OneDrive%20-%20Escuela%20Superior%20de%20Administraci%C3%B3n%20Publica/Antigravity/ConversorPDFv1.6.3.72/.agents/skills/swarm-planner/references/plan-contract.md).

```markdown
# Plan: [Task Name]
**Generated**: [Date] · **Author**: Cline/Swarm-Planner · **Estado**: draft

## Execution envelope
- branch: refactor/<slug>
- workspace: <copia del arbol real, validada con scripts/preflight.ps1>
- baseline: tests/baseline/test_baseline.txt (total: 51, passed: 50, failed: 1)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0: [ConversorPDF.XLL.csproj, Core/Helpers/App.cs, **/Migrations/**, SchemaMigrator.cs, RibbonController.cs, composition root]
- stop_conditions: [destructivo, seguridad, push/merge/publish, plan irrecuperable tras 3 fixes]
- ledger_owner: swarm-parallel-task
- no_push: true

## Overview
[Summary of task and approach]

## Dependency Graph
```
[Visual representation of task dependencies]
```

## Tasks

### T1: [Name]
- **depends_on**: []
- **location**: [file paths]
- **files_lock**: [specific files to touch]
- **description**: [what to do]
- **validation**: [acceptance criteria]
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "..."
- **evidence**: [to be filled by build-verifier]
- **status**: Not Completed
- **log**: [leave empty, to be filled out later]
- **files edited/created**: [leave empty, to be filled out later]

### T2: [Name]
- **depends_on**: []
- **location**: [file paths]
- **description**: [what to do]
- **validation**: [how to verify]
- **status**: Not Completed
- **log**: [leave empty, to be filled out later]
- **files edited/created**: [leave empty, to be filled out later]

### T3: [Name]
- **depends_on**: [T1]
- **location**: [file paths]
- **description**: [what to do]
- **validation**: [how to verify]
- **status**: Not Completed
- **log**: [leave empty, to be filled out later]
- **files edited/created**: [leave empty, to be filled out later]

[... continue for all tasks ...]

## Parallel Execution Groups

| Wave | Tasks | Can Start When |
|------|-------|----------------|
| 1 | T1, T2 | Immediately |
| 2 | T3, T4 | Wave 1 complete |
| 3 | T5 | T3, T4 complete |
| ... | ... | ... |

## Testing Strategy
- [How to test]
- [What to verify]

## Risks & Mitigations
- [What could go wrong + how to handle]
```

## Important

- Every task must have explicit `depends_on` field
- Root tasks (no dependencies) can be executed in parallel immediately
- Do NOT implement - only create the plan
- Always use Context7 for external dependencies before finalizing tasks
- Always ask questions where ambiguity exists
