# PLAN CONSOLIDADO — Arquitectura multiagente (modelo Superpowers) para ConversorPDF

> **Documento para implementación por otro modelo.** Es autocontenido: incluye verdad verificada del entorno, diseño, especificación de cada archivo, tareas atómicas con dependencias, criterios de aceptación y el protocolo de tres fases (Cline planifica → Antigravity ejecuta → Cline cierra).
> Guardar como: `.agents/plans/2026-10-27-multiagent-superpowers-plan.md`
> Raíz del repo: `C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`

---

## 0. Instrucciones para el modelo implementador

1. **No re-investigues** lo que está en §1: fue verificado en disco el 2026-10-27. Si algo contradice §1, **detente y reporta**.
2. Implementa en el orden de §7 (olas). Cada tarea se considera cerrada **solo** con la validación de §6 ejecutada y su salida pegada.
3. **Todo es aditivo** salvo 3 archivos existentes (`test-convention.md`, `development-process/SKILL.md`, `swarm-planner/SKILL.md`) → haz commit de respaldo antes (T0).
4. Antes de escribir, lee `.agents/rules/*.md` completos (7 archivos). Son `always_on` y no son negociables.
5. **Prohibido** (§10): migrar el `.csproj` a SDK-style, migrar Moq→NSubstitute, dividir el proyecto de tests, cambiar versión, hacer push/merge.

---

## 1. Verdad verificada del entorno (2026-10-27)

| Hecho | Evidencia |
|---|---|
| **No hay `git` en PATH** | `git --version` → comando no encontrado |
| Sí existe repo Git | `.git/HEAD → refs/heads/refactor/academic-modularization`; remoto `https://github.com/fexluis/COM-PDF-Merge.git`; `user.email=fexluis@gmail.com` (sin `user.name`) |
| Toolchain disponible | `dotnet 10.0.301`, MSBuild `18.6.4+96856fd72`, `C:\Program Files\dotnet\sdk\10.0.301\vstest.console.dll`, reference assemblies `.NETFramework\v4.8`; runtime `.NET 4.8.10008` (Release 534296) |
| **No hay Visual Studio instalado** | `vswhere.exe` ausente; `"C:\Program Files*\Microsoft Visual Studio"` no existe |
| Build por CLI ya usado | `test_output.txt` muestra warnings `MSB4011` del SDK de dotnet sobre `ConversorPDF.XLL.csproj` |
| **No hay `.sln`** | solo `ConversorPDF.XLL.csproj` + `ConversorPDF.Tests.csproj` |
| XLL es **legacy no-SDK** | 451 entradas `<Compile Include>` vs 452 `.cs`; `PackageReference` mezclado con `<Reference><HintPath>`; `Platform` por defecto `x64`; `Version/AssemblyVersion 1.6.3.79`; ExcelDna 1.8.0 (`RunExcelDnaPack=false`, `ExcelDnaUnpack=true`) |
| Tests es SDK-style con globbing | `net48`, `UseWindowsForms`, xunit 2.7.0, xunit.runner.visualstudio 2.5.7, **FluentAssertions 6.12.0**, **Moq 4.20.70**, MSTest 3.3.1, Dapper 2.1.35, System.Data.SQLite 1.0.119, SQLitePCLRaw.core 2.1.8; `ProjectReference` → XLL |
| 138 archivos de test | carpetas `Academic/{Application,Domain}`, `Core`, `Integration`, `UI`, `Unit` |
| Migraciones versionadas | `SchemaMigrator` ordena `IEnumerable<IMigration>` por `Version`; existen `MigrationV2…V47` |
| **Versión duplicada** | `Core/Helpers/App.cs:72 APP_VERSION="1.6.3.79"` + `:14-15 [AssemblyVersion]/[AssemblyFileVersion]` + `csproj Version` |
| Composition root no localizado | grep `AddSingleton|AddTransient|AddScoped|ServiceCollection` en XLL → **0 resultados** (registro manual o en otro archivo) |
| `.gitignore` incompleto | ignora `/ConversorPDF/bin|obj` (ruta inexistente); **no** ignora `ConversorPDF.XLL/bin|obj`, `ConversorPDF.Tests/bin|obj`, `TestResults/` (≈472 archivos binarios presentes) |
| Skill/skills actuales | `.agents/{rules(7), skills(3: development-process, swarm-planner, swarm-parallel-task), workflows(7), checkpoints}`; **no existe `.agents/agents/`** |
| Deriva regla↔realidad | `test-convention.md` exige NSubstitute + FluentAssertions 7+ + proyectos `*.Tests.Unit` por capa; realidad = Moq + FA 6.12 + proyecto único |
| Regla incompatible con agentes | `development-process/SKILL.md`: *"Do not use CLI `dotnet test`… use Visual Studio Test Explorer"* |
| Ruta dentro de **OneDrive** | locks/archivos "en uso" con MSBuild/VSTest; worktrees inviables |
| Harness Cline actual | **sin** `invoke_subagent` → no puede ejecutar subagentes; sí puede leer todo el repo e invocar `swarm-planner` |

**Superficies Antigravity (docs oficiales):** skills en `<workspace>/.agents/skills/<skill>/`; subagentes custom en `.agents/agents/<name>.md` (o `<name>/agent.md`) con frontmatter `name, description, tools[], mainAgent, subagent, model (inherit|flash|pro), commandExecutionPolicy (off|auto|eager|sandbox), mcpServers, skills/plugins`; invocación vía `invoke_subagent`; opciones de workspace `inherit|branch|share`; **el subagente no hereda contexto**; built-ins `research`, `browser`, `self`; panel `/agents`; `Alt+J` (saltar a subagente pendiente), `Ctrl+K` (aprobar); CLI convierte cada skill en slash command; `agy plugin install <url>`.

---

## 2. Objetivo y no-objetivos

**Objetivo.** Que una feature o refactor se ejecute con: plan verificable escrito por Cline → ejecución en olas por subagentes Antigravity con evidencia RED→GREEN → cierre independiente por Cline. Con **un solo escritor por archivo**, **un solo dueño de archivos calientes** y **un único comando de verificación**.

**No-objetivos (explícitos):** migrar el `.csproj` a SDK-style · cambiar frameworks de test/mocking · dividir `ConversorPDF.Tests` por capas · instalar el plugin `obra/superpowers` (se toma el *modelo*, no el paquete) · automatizar push/PR.

---

## 3. Diseño

```
┌──────────── PLANO DE CONTROL · Cline ────────────┐
│ A1 context pack  A2 /swarm-planner  A3 review    │
│ A4 handoff (.md)  A5 aprobación  A6 pre-flight   │
└───────────────┬─────────────────────────────────┘
                │  .agents/plans/<slug>-plan.md  +  <slug>-handoff.md
                ▼
┌──────────── PLANO DE EJECUCIÓN · Antigravity App/CLI ───────────┐
│ B0 plan-reviewer gate → B1 olas → B2 tdd-red-writer (paralelo)  │
│ B3 implementer (paralelo) → B4 integrator (L0 serializado)      │
│ B5 build-verifier (verify.ps1) → task-reviewer → final review   │
│ ledger en el plan + commit T<n> (nunca push)                    │
└───────────────┬─────────────────────────────────────────────────┘
                │  plan actualizado (ledger) + commits + _evidence/
                ▼
┌──────────── PLANO DE CIERRE · Cline ────────────────────────────┐
│ C1 relee ledger+git  C2 re-ejecuta FULL  C3 grep invariantes    │
│ C4 walkthrough + docs/articles/changes  C5 cierra plan          │
└─────────────────────────────────────────────────────────────────┘
```

**Roles (6 subagentes).**

| Rol | Escribe | Responsabilidad |
|---|---|---|
| `plan-reviewer` | ❌ | Dependencias faltantes, orden, criterios no verificables, solapes de lease |
| `tdd-red-writer` | solo `ConversorPDF.Tests/**` | Tests que fallan por la razón correcta + salida pegada |
| `implementer` | solo su lease | GREEN mínimo, sin debilitar tests, sin tocar L0 |
| `integrator` | **L0** | `csproj`, `APP_VERSION`, migraciones, `RibbonController`, composition root |
| `task-reviewer` | ❌ | Spec compliance + calidad (YAGNI, hygiene de tests, invariantes) |
| `build-verifier` | solo `_evidence/` | Ejecuta `verify.ps1`, produce TRX+resumen, decide RED/GREEN |

**Leases.**
- **L0 — exclusivo global, solo `integrator`, 1 a la vez:** `ConversorPDF.XLL.csproj` · `Core/Helpers/App.cs` · `**/Migrations/MigrationV*.cs` · `SchemaMigrator.cs` · `Ribbon/RibbonController.cs` · composition root · `docs/toc.yml`.
- **L1 — por archivo o por slice:** `Academic/<Slice>/**` (preferir lease por archivo cuando dos tareas comparten carpeta).
- **L2 — solo lectura:** `.agents/**`, plan, handoff, `docs/**` (excepto `docs/articles/changes/`, que es del cierre).

**Gate de verificación:** `scripts/verify.ps1` (Build/Fast/Full) + baseline. **Nada se declara verde sin TRX.**

**Ledger:** un solo escritor (el orquestador). Línea por tarea + `Ruling: <decisión> — <porqué> — <coste si me equivoco>`.

---

## 4. Inventario de artefactos

| Ruta | Acción | Dueño | Tarea |
|---|---|---|---|
| `.gitignore` | modificar | Cline | T2 |
| `docs/architecture/build-verification.md` | crear | Cline | T4 |
| `scripts/verify.ps1` | crear | Cline | T5 |
| `tests/baseline/test_baseline.txt` / `.trx` | crear | Cline | T6 |
| `.agents/rules/parallel-work.md` | crear | Cline | T7 |
| `.agents/rules/verification-commands.md` | crear | Cline | T8 |
| `.agents/rules/test-convention.md` | modificar | Cline | T9 |
| `.agents/skills/development-process/SKILL.md` | modificar | Cline | T10 |
| `.agents/skills/swarm-planner/references/plan-contract.md` | crear | Cline | T11 |
| `.agents/skills/swarm-planner/SKILL.md` | modificar | Cline | T11 |
| `.agents/plans/README.md`, `.agents/context/README.md` | crear | Cline | T12 |
| `.agents/agents/{plan-reviewer,tdd-red-writer,implementer,integrator,task-reviewer,build-verifier}.md` | crear | Cline | T13–T18 |
| `.agents/skills/swarm-parallel-task/SKILL.md` | reescribir | Cline | T20 |
| `.agents/skills/executing-plans-local/SKILL.md` | crear | Cline | T21 |
| `scripts/context-pack.ps1` + `.agents/context/_template.md` | crear | Cline | T23 |
| `.agents/context/academic-genetica.md` | crear | Cline | T24 |
| `.agents/plans/<slug>-plan.md` / `-handoff.md` | crear | Cline | T25 |
| `_evidence/**` | crear (runtime) | build-verifier | B5 |
| `docs/articles/changes/YYYY-MM-DD-*.md`, `walkthrough.md` | crear | Cline | C4 |

---

## 5. Especificación de archivos

### 5.1 `scripts/verify.ps1` (propuesta; **calibrar en T4**)

Contrato: `-Mode Build|Fast|Full`, `-Filter`, `-Configuration Debug`, `-Platform x64`, `-NoRestore`, `-EvidenceRoot _evidence`.
Salidas: `_evidence/<yyyyMMdd-HHmmss>/{build.log,test.trx,summary.md}`.
Exit codes: **0** ok · **1** build falló · **2** tests fallaron · **3** regresión vs baseline · **4** tooling ausente.

```powershell
[CmdletBinding()]
param(
    [ValidateSet('Build','Fast','Full')][string]$Mode = 'Fast',
    [string]$Filter,
    [string]$Configuration = 'Debug',
    [string]$Platform = 'x64',
    [switch]$NoRestore,
    [string]$EvidenceRoot = '_evidence'
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$xll  = Join-Path $repo 'ConversorPDF.XLL\ConversorPDF.XLL.csproj'
$tst  = Join-Path $repo 'ConversorPDF.Tests\ConversorPDF.Tests.csproj'
$ev   = Join-Path $repo (Join-Path $EvidenceRoot (Get-Date -Format 'yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Force -Path $ev | Out-Null

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) { Write-Error 'dotnet no encontrado'; exit 4 }
if ($repo -match 'OneDrive') { Write-Warning 'Repo en OneDrive: riesgo de locks. Usa la copia local.' }

$restore = if ($NoRestore) { '--no-restore' } else { '' }
$buildLog = Join-Path $ev 'build.log'

Write-Host "[1/3] Build XLL ($Platform/$Configuration)" -ForegroundColor Cyan
& dotnet build $xll -c $Configuration -p:Platform=$Platform $restore -v:m 2>&1 | Tee-Object -FilePath $buildLog
if ($LASTEXITCODE -ne 0) { Write-Error 'Build XLL falló'; exit 1 }

Write-Host "[2/3] Build tests" -ForegroundColor Cyan
& dotnet build $tst -c $Configuration -p:Platform=$Platform $restore --no-dependencies -v:m 2>&1 |
    Tee-Object -FilePath $buildLog -Append
if ($LASTEXITCODE -ne 0) { Write-Error 'Build tests falló'; exit 1 }
if ($Mode -eq 'Build') { exit 0 }

Write-Host "[3/3] Tests ($Mode)" -ForegroundColor Cyan
$trx = Join-Path $ev 'test.trx'
$args = @($tst, '-c', $Configuration, '--no-build', '--logger', "trx;LogFileName=$trx",
          '--results-directory', $ev)
if ($Filter) { $args += @('--filter', $Filter) }
& dotnet test @args 2>&1 | Tee-Object -FilePath (Join-Path $ev 'test.log')
$testExit = $LASTEXITCODE

if (-not (Test-Path $trx)) { Write-Error 'No se generó TRX'; exit 2 }
[xml]$doc = Get-Content -LiteralPath $trx
$results = $doc.SelectNodes("//*[local-name()='UnitTestResult']")
$failed  = @($results | Where-Object { $_.outcome -eq 'Failed' } | ForEach-Object { $_.testName })
$total   = $results.Count
$passed  = @($results | Where-Object { $_.outcome -eq 'Passed' }).Count

$baselineFile = Join-Path $repo 'tests\baseline\test_baseline.txt'
$baselineFailures = @()
if (Test-Path $baselineFile) {
    $baselineFailures = Get-Content $baselineFile | Where-Object { $_ -like 'FAILED *' } |
                        ForEach-Object { $_.Substring(7) }
}
$regressions = @($failed | Where-Object { $baselineFailures -notcontains $_ })

@(
  "# Verificación $Mode — $(Get-Date -Format 'u')"
  ""
  "- Comando: dotnet test ($Configuration/$Platform)"
  "- Filtro: $(if ($Filter) { $Filter } else { '<todos>' })"
  "- Total: $total · Passed: $passed · Failed: $($failed.Count)"
  "- Baseline: $($baselineFailures.Count) fallos conocidos"
  "- Regresiones: $($regressions.Count)"
  ""
  "## Fallos"
  $(if ($failed.Count) { $failed | ForEach-Object { "- $_" } } else { "- ninguno" })
  ""
  "## Regresiones (no estaban en baseline)"
  $(if ($regressions.Count) { $regressions | ForEach-Object { "- $_" } } else { "- ninguna" })
) | Set-Content -LiteralPath (Join-Path $ev 'summary.md') -Encoding UTF8

Write-Host "Evidencia: $ev" -ForegroundColor Green
if ($Mode -eq 'Full' -and $regressions.Count -gt 0) { exit 3 }
if ($testExit -ne 0) { exit 2 }
exit 0
```

> **T4 debe decidir** (no asumir): si el test host corre x64, si `--no-dependencies` es necesario, si los tests `UI/` (MSTest STA) requieren filtro (`--filter "TestCategory!=UI"` o equivalente) y si `Platform=AnyCPU` funciona. Congela la combinación ganadora en `docs/architecture/build-verification.md` y en `verification-commands.md`.

### 5.2 `.agents/rules/parallel-work.md` (nuevo, `trigger: always_on`)

Contenido mínimo obligatorio:
- Tabla de roles y permisos de escritura (§3).
- Lista L0 textual de archivos calientes (copiar de §3) + "solo `integrator`, una tarea a la vez".
- Regla: **dos tareas no pueden solaparse en `files_lock`**; el orquestador debe **rechazar o serializar** (log obligatorio).
- `max_parallel_writers: 2`; revisores de solo lectura sin límite.
- Regla C#: **el RED debe compilar** → el esqueleto del tipo nuevo lo crea el `integrator` **antes** que los tests.
- Regla de migraciones: **nunca** elegir `V<n>` en paralelo; el `integrator` asigna el siguiente número en el momento de escribir, leyendo `Migrations/` y `__SchemaVersion`.
- Regla de versión: `APP_VERSION` y `csproj Version` solo `integrator`, solo al cerrar release.
- Regla de plan: un solo escritor del plan (orquestador); Cline lo trata como read-only tras el handoff.

### 5.3 `.agents/rules/verification-commands.md` (nuevo)

Comandos congelados + contrato de evidencia:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~<Clase>"
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
```
Reglas: RED = test nuevo/actualizado fallando **por la razón esperada** (pegar nombre + mensaje del TRX). GREEN = mismo test Passing **y** `Full` sin regresiones vs baseline. Prohibido: `--filter` que excluya el test propio, debilitar asserts, subir timeouts, reordenar tests para ocultar flakiness.

### 5.4 Parches a reglas existentes

**`.agents/rules/test-convention.md`** — sustituir el bloque "Frameworks & Tools" y "Naming & Organization" por la realidad, sin romper el resto:
```
- Framework: xUnit 2.7 (principal). MSTest 3.3.1 solo para UI/STA (carpeta UI/).
- Mocking: Moq 4.20.70 (NSubstitute NO está instalado; no introducirlo sin ADR).
- Assertions: FluentAssertions 6.12.0 (la sintaxis de v7+ no está disponible).
- Proyecto: único ConversorPDF.Tests (net48) con carpetas espejo del XLL
  (Academic/Application, Academic/Domain, Core, Integration, UI, Unit).
- Archivo: {SystemUnderTest}Tests.cs ; Método: {Accion}_When_{Condicion}_Should_{Resultado}
```
Añadir nota: *"Esta regla refleja el estado actual; no migrar frameworks dentro de tareas de feature."*

**`.agents/skills/development-process/SKILL.md`** — reemplazar el paso 4 ("Visual Studio Test Explorer") por:
```
- Ejecutar `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~<Clase>"` durante el ciclo
  y `-Mode Full` al cerrar cada sub-tarea. La evidencia (TRX + summary.md) va al plan.
- VS Test Explorer solo como verificación humana opcional.
```

### 5.5 `.agents/skills/swarm-planner/references/plan-contract.md` (nuevo)

Debe contener: el **Execution Envelope** (branch, workspace, baseline, verify, max_parallel_writers, agents, hot_paths_L0, stop_conditions, ledger_owner, no_push) · los campos obligatorios de tarea (`id, depends_on, location, files_lock, description, validation, verification, evidence, status, log, files edited/created, ruling`) · el formato de ledger · la política de `Ruling` · el tope de rondas de fix (3) · la regla "el plan se escribe en `.agents/plans/`, nunca en la raíz".

### 5.6 Los 6 subagentes (`.agents/agents/*.md`)

| Archivo | `model` | `commandExecutionPolicy` | `tools` (documentados) |
|---|---|---|---|
| `plan-reviewer.md` | `pro` | `off` | `view_file`, `grep_search` |
| `tdd-red-writer.md` | `inherit` | `sandbox` | `view_file`, `grep_search`, `replace_file_content`, `run_command` |
| `implementer.md` | `pro` | `sandbox` | `view_file`, `grep_search`, `replace_file_content`, `run_command` |
| `integrator.md` | `pro` | `auto` | `view_file`, `grep_search`, `replace_file_content`, `run_command` |
| `task-reviewer.md` | `pro` | `off` | `view_file`, `grep_search` |
| `build-verifier.md` | `flash` | `auto` | `view_file`, `run_command` |

Plantilla (ejemplo `implementer.md`):

```markdown
---
name: implementer
description: Implements exactly one task from a swarm plan in the ConversorPDF repo, until its tests are green. Use when a task is unblocked and its files_lock is free.
subagent: true
mainAgent: false
model: pro
commandExecutionPolicy: sandbox
tools: [view_file, grep_search, replace_file_content, run_command]
skills: [development-process]
---
# Rol
Implementas UNA tarea (T<n>) del plan indicado. No eres el orquestador.

# Reglas duras (leer antes de escribir)
- `.agents/rules/architecture-rules.md`, `desktop-code.md`, `test-convention.md`,
  `single-type-per-file.md`, `sqlite-excel-dna-gotchas.md`, `parallel-work.md`.
- PROHIBIDO tocar: ConversorPDF.XLL.csproj, Core/Helpers/App.cs, **/Migrations/**,
  SchemaMigrator.cs, Ribbon/RibbonController.cs  → son L0 del `integrator`.
- PROHIBIDO: C# 12 (`[]`, primary ctors), MediatR, IOptions, Interop/WinForms/DataGridView
  en Application, DateTime.Now (usa IDateTimeProvider), .Result/.Wait().
- PROHIBIDO debilitar, saltar o borrar tests. Los tests de la fase RED son la especificación.
- Result Pattern: `Result`/`Result<T>`, sin excepciones para reglas de negocio.

# Procedimiento
1. Lee el plan y la tarea completa. Lee los archivos del `location` (y solo esos + 1 nivel de contexto).
2. Implementa el mínimo cambio que satisface los criterios de aceptación.
3. Verifica: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~<Clase>"`
4. Si tocas más archivos de los declarados en `files_lock`: PARA y reporta BLOCKED.
5. Devuelve: archivos, cambio, evidencia RED→GREEN (ruta del summary.md), riesgos. Sin push.
```

Los otros 5 siguen el mismo esqueleto cambiando: rol, permisos de escritura (`tdd-red-writer` solo `ConversorPDF.Tests/**`; `integrator` solo L0; revisores sin escritura), y contrato de salida (`plan-reviewer` → hallazgos numerados con severidad; `task-reviewer` → veredicto `OK | CAMBIOS_REQUERIDOS` con hallazgos `Critical|Important|Minor`; `build-verifier` → ruta de `_evidence/…/summary.md` + exit code).

### 5.7 Plantillas de plan y handoff

`plan-template.md` (encabezado obligatorio):

```markdown
# Plan: <título>
**Generado**: <fecha> · **Autor**: Cline/Swarm-Planner · **Estado**: draft|approved|executing|closed

## Execution envelope
- branch: refactor/<slug> · workspace: C:\dev\ConversorPDF
- baseline: tests/baseline/test_baseline.txt (total/pass/fail/skip = <N>)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0: [...]   · stop_conditions: destructivo | seguridad | push/merge/publish | plan irrecuperable
- ledger_owner: swarm-parallel-task · no_push: true

## Dependency graph
## Tasks (T1..Tn con id/depends_on/location/files_lock/description/validation/verification/evidence/status/log)
## Parallel execution groups (Wave | Tasks | Can start when)
## Testing strategy · Risks & mitigations
```

`handoff-template.md` (una página, 7 puntos): workspace a abrir · rama y estado limpio · comando exacto `/swarm-parallel-task <plan>` · gate obligatorio `plan-reviewer` · criterio de aceptación · prohibiciones · salida esperada (plan actualizado + resumen).

### 5.8 `swarm-parallel-task` (reescritura) — puntos obligatorios

1. Parsear envelope + tareas; **abortar** si falta baseline, `verify.ps1` o `plan-reviewer`.
2. Gate: invocar `plan-reviewer`; si hay `Critical` → `BLOCKED`, no ejecutar.
3. Calcular olas; **detección de solape de `files_lock`** (prefijo/ruta) → serializar y loguear `LEASE-SERIALIZED T4←T5`.
4. Por ola: `tdd-red-writer` (paralelo, leases disjuntos) → `implementer` (paralelo) → `integrator` (serial) → `build-verifier` → `task-reviewer` (2 etapas: spec → calidad).
5. `integrator-first` cuando la tarea introduce tipos nuevos (esqueleto + `<Compile Include>` antes del RED).
6. Tope de 3 rondas de fix; luego `Ruling` en el ledger y seguir (o `BLOCKED` si es irreparable).
7. Commit por tarea **si hay git**: `git add -- <archivos del lease>` + `git commit -m "T<n>: <resumen>"`. **Nunca push**.
8. Actualizar ledger en el mismo paso que el commit.
9. Cierre: `final review` de rama + `verify.ps1 -Mode Full`.
10. Fallback documentado: si `invoke_subagent` no existe → `executing-plans-local` (secuencial, mismo ledger y misma evidencia).

---

## 6. Tareas (formato de ejecución)

| ID | depends_on | Tarea | location | Validación |
|---|---|---|---|---|
| **T0** | `[]` | Commit de respaldo del estado actual (si se instala git) o copia `.bak` de los 3 archivos a modificar | `.agents/**` | `git log -1` / `.bak` existen |
| **T1** | `[]` | Instalar Git CLI + `user.name`; verificar repo | máquina | `git --version`; `git status`; `git rev-parse --abbrev-ref HEAD` = `refactor/academic-modularization` |
| **T2** | `[T1]` | `.gitignore` completo (`ConversorPDF.XLL/bin|obj`, `ConversorPDF.Tests/bin|obj`, `TestResults/`, `_evidence/`, `*.user`, `.vs/`, `obj/`, `_FIXES/`) + destrackear binarios si están en índice | `.gitignore` | `git status --porcelain` sin `bin/`/`obj/`/`TestResults`; `git ls-files | Select-String 'bin|obj'` vacío |
| **T3** | `[T1,T2]` | Copia local fuera de OneDrive (`C:\dev\ConversorPDF`) + regla de sincronización documentada | `docs/architecture/` | build+tests en la copia, misma salida que en OneDrive |
| **T4** | `[T3]` | **Calibrar** build/test CLI (matriz de comandos) | `docs/architecture/build-verification.md` | Registro con comando ganador, plataforma del test host, filtro UI, y 2 ejecuciones reproducibles |
| **T5** | `[T4]` | `scripts/verify.ps1` | `scripts/verify.ps1` | `-Mode Build` exit 0; `-Mode Fast -Filter` exit 0; `-Mode Full` genera TRX+summary |
| **T6** | `[T5]` | Baseline de la suite | `tests/baseline/` | `test_baseline.txt` con total/pass/fail + `FAILED <test>`; 2 corridas con los mismos números |
| **T7** | `[T6]` | `parallel-work.md` (leases L0/L1/L2, límites, reglas C#/migración/versión) | `.agents/rules/parallel-work.md` | Revisión cruzada: toda ruta L0 de §3 aparece; un plan de prueba con solape es rechazado/serializado |
| **T8** | `[T5]` | `verification-commands.md` (comandos congelados + contrato de evidencia) | `.agents/rules/verification-commands.md` | Los 3 comandos funcionan tal cual |
| **T9** | `[T6]` | Alinear `test-convention.md` (Moq/FA6.12/proyecto único/xUnit+MSTest) | `.agents/rules/test-convention.md` | Un test generado solo con lo permitido compila y pasa |
| **T10** | `[T5]` | Parchear `development-process` (CLI vía verify.ps1) | `.agents/skills/development-process/SKILL.md` | Skill cita comandos existentes; ejemplo ejecutado |
| **T11** | `[T7]` | `plan-contract.md` + parche `swarm-planner` (envelope, leases, salida a `.agents/plans/`) | `.agents/skills/swarm-planner/**` | Plan generado contiene envelope y los 11 campos por tarea |
| **T12** | `[T11]` | `README.md` de `.agents/plans/` y `.agents/context/` (plantillas + ownership) | `.agents/plans/README.md`, `.agents/context/README.md` | Plantillas enlazadas desde el README y usadas en T25 |
| **T13** | `[T11]` | `plan-reviewer.md` | `.agents/agents/` | Frontmatter válido; detecta 2 huecos sembrados a propósito |
| **T14** | `[T11]` | `tdd-red-writer.md` | `.agents/agents/` | En prueba: solo escribe en `ConversorPDF.Tests/**` |
| **T15** | `[T11]` | `implementer.md` | `.agents/agents/` | Rechaza tarea con L0 en `files_lock` (reporta BLOCKED) |
| **T16** | `[T11]` | `integrator.md` | `.agents/agents/` | Añade `<Compile Include>` correctamente y sin duplicar (CS2002) |
| **T17** | `[T11]` | `task-reviewer.md` | `.agents/agents/` | Veredicto con severidades sobre un diff real |
| **T18** | `[T11]` | `build-verifier.md` | `.agents/agents/` | Devuelve ruta de `summary.md` + exit code |
| **T19** | `[T13..T18]` | Smoke test de subagentes en Antigravity App/CLI (descubrimiento + tool names) | sesión Antigravity | Los 6 aparecen en `/agents`; ninguno cuelga; corregir nombres si aplica |
| **T20** | `[T8,T19]` | Reescribir `swarm-parallel-task` (§5.8) | `.agents/skills/swarm-parallel-task/SKILL.md` | Plan de 4 tareas en 2 olas: 0 archivos fuera de lease; ledger completo |
| **T21** | `[T20]` | `executing-plans-local` (fallback sin `invoke_subagent`) | `.agents/skills/executing-plans-local/SKILL.md` | Mismo plan ejecutado secuencialmente con evidencia equivalente |
| **T22** | `[T11]` | *(opcional)* `verification-before-completion` + `systematic-debugging` adaptados a COM/Excel-DNA | `.agents/skills/*` | Caso real diagnosticado con evidencia |
| **T23** | `[T12]` | `scripts/context-pack.ps1` (greps de §A1) + `_template.md` | `scripts/`, `.agents/context/` | Genera pack ≤400 líneas de un slice |
| **T24** | `[T23]` | Pack real del slice `Genetica` | `.agents/context/academic-genetica.md` | 6 secciones completas; ≤400 líneas |
| **T25** | `[T3..T6,T19,T20,T24]` | **Piloto** `ag-penalizaciones` (6 tareas, 5 olas) en Antigravity | `.agents/plans/2026-10-27-ag-penalizaciones-plan.md` | AC1–AC5 de §9 |
| **T26** | `[T25]` | Post-mortem + ajuste de reglas con datos (olas, conflictos, rondas) | `.agents/plans/`, `.agents/rules/parallel-work.md` | Cambios justificados por métricas del ledger |

---

## 7. Olas de ejecución

| Ola | Tareas | Puede empezar cuando |
|---|---|---|
| 0 | T0 → T1 | inmediato |
| 1 | T2 → T3 → T4 → T5 → T6 | serial (cadena de dependencias) |
| 2 | **T7, T8, T9, T10** | T6 completo (T8 tras T5) |
| 3 | **T11** → T12 · **T13, T14, T15, T16, T17, T18** (paralelo) | T7/T11 |
| 4 | T19 → T20 → T21 · T23 → T24 (T22 opcional) | T19 / T12 |
| 5 | T25 | T24, T20, T6 |
| 6 | T26 | T25 |

Regla: en la ola 3 los 6 archivos de agentes son **independientes** → paralelizables; nadie toca L0.

---

## 8. Protocolo operativo (A → B → C)

### FASE A — Cline planifica

```text
Prompt a Cline (Act mode):
"Planifica <feature>. Usa /swarm-planner con el context pack .agents/context/<slice>.md.
Escribe .agents/plans/<fecha>-<slug>-plan.md y -handoff.md según .agents/plans/README.md.
Antes de ceder: autoevaluación con el checklist de 12 puntos y corrige el plan."
```

**Checklist de 12 puntos (Cline, determinista):** 1 dependencias completas · 2 orden ejecutable · 3 ninguna tarea mezcla L0 sin `integrator` · 4 `files_lock` disjuntos por ola · 5 cada `verification` es un comando existente · 6 cada `validation` es falsable · 7 el RED compila (esqueleto previsto) · 8 migración/versión asignadas por `integrator` · 9 no hay tareas >1 slice · 10 presupuesto de contexto del pack ≤400 líneas · 11 criterio de aceptación global definido · 12 riesgos con mitigación.

Handoff: **2 archivos** (`-plan.md` + `-handoff.md`). Aprobación humana. Pre-flight: `git status --porcelain` vacío, `Test-Path` de `verify.ps1`, baseline, agentes.

### FASE B — Antigravity ejecuta

```text
B0  Abrir workspace C:\dev\ConversorPDF   (no OneDrive)
B1  /swarm-parallel-task .agents/plans/2026-10-27-ag-penalizaciones-plan.md
B2  gate plan-reviewer → si Critical: BLOCKED
B3  olas: tdd-red-writer (paralelo) → implementer (paralelo) → integrator (serial L0)
B4  build-verifier: verify.ps1 -Mode Full  ·  task-reviewer (spec → calidad)  ·  final review
B5  ledger + commit "T<n>: ..." (sin push)
```

`workspace`: `inherit` por defecto; `branch` (worktree) solo si hay git y dos tareas comparten carpeta; `share` nunca (rompe el aislamiento).

### FASE C — Cline cierra (no acepta el ledger)

```powershell
Get-Content .agents\plans\2026-10-27-ag-penalizaciones-plan.md      # ledger
git log --oneline -n 12; git diff --stat <base>..HEAD
powershell -ExecutionPolicy Bypass -File scripts\verify.ps1 -Mode Full                             # Cline re-ejecuta
Select-String -Path ConversorPDF.XLL\Academic\Application\**\*.cs -Pattern 'Excel\.|Microsoft\.Office|DataGridView'
Select-String -Path ConversorPDF.XLL\Academic\**\*.cs -Pattern 'new \[\]|DateTime\.Now|\.Result\b|\.Wait\(\)'
```
Luego escribe `walkthrough.md` + `docs/articles/changes/YYYY-MM-DD-<slug>.md` (exigido por `development-process`) y marca el plan `closed`. Si hay hallazgos → **plan nuevo** `T-fix`; nunca reabrir el plan original con dos escritores.

---

## 9. Definition of Done y pruebas de aceptación

**DoD de la arquitectura:** §4 completo; los 6 agentes descubiertos; `verify.ps1` único camino de verificación; reglas sin deriva; `.gitignore` limpio; copia local operativa.

| AC | Prueba | Esperado |
|---|---|---|
| **AC1** | `verify.ps1 -Mode Full` en estado limpio | exit 0 y summary con los números del baseline |
| **AC2** | Sembrar un assert falso | exit 2 y el test aparece en `summary.md`; revertir |
| **AC3** | Plan de prueba con dos `files_lock` solapados | el orquestador serializa (log `LEASE-SERIALIZED`) o rechaza |
| **AC4** | Piloto (6 tareas, 5 olas) | RED→GREEN por tarea; 0 archivos fuera de lease; solo `integrator` tocó L0; ledger completo |
| **AC5** | Cierre Cline | re-ejecución Full coincide con el ledger; grep de invariantes limpio; `git log origin/<branch>` sin cambios (no push) |

---

## 10. Riesgos, rollback y prohibiciones

| Riesgo | Mitigación |
|---|---|
| `.csproj` con 451 `Compile` → conflictos | L0 exclusivo + `integrator-first` + lease por archivo |
| Migraciones `V<n>` colisionando | número asignado solo por `integrator`, leyendo `Migrations/` en el momento |
| Tests que "se arreglan" debilitándolos | `task-reviewer` + regla dura + diff de tests revisado |
| OneDrive (locks, corrupción) | copia local `C:\dev` obligatoria para swarms |
| Sin git → sin commits/worktrees | T1 bloqueante; fallback `executing-plans-local` con ledger |
| Nombre de tool mal escrito cuelga el subagente (limitación conocida) | allowlist mínima + T19 |
| Subagentes solo en App/CLI (no IDE) | verificar en T19 antes de invertir |
| Contexto inflado (452 archivos) | context pack ≤400 líneas + lectura acotada al `location` |
| Baseline desconocido → "verde" falso | T6 antes de cualquier ola de ejecución |

**Rollback:** todo es aditivo; los 3 parches se revierten con el commit de T0. Nada toca producción de la app.

**Prohibiciones:** push/merge · `git commit --amend` sobre commits de otras tareas · migrar `.csproj` a SDK-style · migrar Moq→NSubstitute o FA 6→7 · dividir `ConversorPDF.Tests` · cambiar `APP_VERSION`/`Version` · ejecutar swarms sobre OneDrive · ejecutar sin baseline · aceptar el ledger sin re-verificar.

---

## 11. Verificaciones abiertas (el implementador **debe** confirmar, no asumir)

1. Flags exactos de `agy` (CLI) para modo headless y para invocar un skill desde terminal, en tu versión instalada.
2. Si tu superficie es **Antigravity IDE**, si descubre `.agents/agents/*.md` (los docs lo documentan para App/CLI).
3. Nombres reales de tools admitidos (`view_file`, `grep_search`, `replace_file_content`, `run_command`) — un nombre inválido cuelga el subagente.
4. Plataforma del test host (`x64` vs `AnyCPU`) y comportamiento de los tests `UI/` (MSTest/STA) en headless.
5. Números reales del baseline (hay `test_failures_*.txt` en la raíz que sugieren fallos preexistentes).
6. Ubicación del composition root / registro de servicios (grep sin resultados de `AddSingleton`).
7. Si `bin/obj` están **trackeados** en el índice (`git ls-files`) → entonces T2 includes `git rm -r --cached` (requiere aprobación humana).

---

### Cómo continuar

Este documento está listo para que lo ejecutes tal cual (T0→T26).
