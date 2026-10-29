---
name: swarm-verify
description: >
  [EXPLICIT INVOCATION ONLY] Cierra un plan swarm (Fase C): re-verifica en disco, audita el código
  de cada Task con detectores de fallos de ejecución, valida leases, numeración (Ronda/Wave ·
  Fase/Task) e invariantes de arquitectura, y sella el plan como closed. Nunca reabre el plan.
metadata:
  invocation: explicit-only
---

# Swarm Verify (Fase C — Cline cierra)

Tercera fase del protocolo A → B → C (ver [`.agents/plans/README.md`](../../plans/README.md) §2). La ejecuta **Cline** una vez que `swarm-parallel-task` (o `executing-plans-local`) terminó la Fase B.

> [!CRITICAL]
> **No aceptas el ledger**: lo re-verificas en disco. Eres el único escritor del plan en esta fase.
> **Nunca ejecutes `git push`**. Si hay hallazgos **no reabres** el plan: creas un plan `T-fix`.

Fuentes de verdad:

- Contrato del plan: [`plan-contract.md`](../swarm-planner/references/plan-contract.md)
- Runbook, plantillas y bloque de cierre: [`plans/README.md`](../../plans/README.md) §5
- Leases y concurrencia: [`parallel-work.md`](../../rules/parallel-work.md)
- **Numeración global (Ronda/Wave · Fase/Task)**: [`plans/REGISTRY.md`](../../plans/REGISTRY.md)

---

## 0. Numeración: `Ronda ≡ Wave`, `Fase ≡ Task`

La revisión por rondas (`.clinerules/workflows/Ronda.md`) y los planes swarm describen el mismo objeto con dos vocabularios. Toda salida de esta fase usa **ambos**, con el número global del registro:

| `Ronda.md` / artefactos de revisión | Plan swarm | Nivel |
|---|---|---|
| `Ronda {n}: {título}` | `Wave {n}` (parallel execution group) | Grupo paralelo (una ola = una ronda) |
| `Fase {n}: {título}` | `T{n}: {Nombre}` | Tarea atómica |
| `| Fase | Severidad | Archivo | Líneas |` | Tabla de hallazgos de `audit-tasks.ps1` | Reporte |

Reglas:

1. El plan trae `numbering:` en el envelope (`ronda_base`, `fase_base`) y cada tarea su campo `- **fase**: <n>`.
2. Si faltan, se resuelven con la fila del plan en [`REGISTRY.md`](../../plans/REGISTRY.md) §3; si el plan es `pre-registro`, se reporta `Fase = T{k}` **sin inventar** números globales.
3. `Wave k` del plan ⇔ `Ronda = ronda_base + k - 1`; `T{k}` ⇔ `Fase = fase_base + k - 1`, salvo que el plan declare `- **fase**` explícito (manda el campo).
4. Cerrar el plan incluye **cerrar su fila** en el registro (§C2.5).

---

## 1. Entrada: ¿es realmente Fase C?

| Requisito | Comprobación |
|---|---|
| Plan en `.agents/plans/<YYYY-MM-DD>-<slug>-plan.md` | `Test-Path` |
| Cabecera `**Estado**: executing` | primera línea del plan |
| Todos los `T<n>` en `status: done` | `Get-Content <plan> \| Select-String 'status:'` |
| Bloque `## Execution envelope` presente | cabecera del plan |
| `base` = SHA anterior a T1 | lo aporta el usuario |

Si el plan está en `draft` (Fase A), `approved` (Fase B pendiente) o ya `closed`, **detente y reporta**: no hay nada que sellar.

## 2. C0 — Precondiciones

```powershell
git status --porcelain                                                                          # debe estar VACIO
powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ws> -Plan <plan>      # exit 0
Test-Path scripts\audit-tasks.ps1                                                               # debe ser True (C1.5)
```

Un working tree sucio invalida el ledger (no se puede atribuir cada cambio a una tarea). En OneDrive `preflight.ps1` devuelve `3`: usar `-AllowOneDrive` sólo con aviso explícito en el acta de cierre.

## 3. C1 — Verificación independiente

```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
```

| exit | Significado | Veredicto |
|---|---|---|
| `0` | OK y `TOTAL >= baseline` | continuar |
| `1` | Build falló | `FIX_PLAN` |
| `2` | Tests fallaron / infraestructura | `FIX_PLAN` |
| `3` | Regresión vs baseline, `SUITE_SHRUNK` o `NO_BASELINE` | `FIX_PLAN` (bloqueante) |
| `4` | Tooling o workspace inválido | `BLOCKED` |

Evidencia: `_evidence/<stamp>/summary.md` — se enlaza por ruta y **no se commitea**.

## 3.1 C1.5 — Auditoría exhaustiva de código por tarea (errores de ejecución)

`verify.ps1` certifica que **compila** y que la suite **pasa**; no ve errores de ejecución silenciosos (deadlocks STA, Interop en capas limpias, SQL destructivo, migraciones duplicadas, tests debilitados, `INSERT` posicional…). Esta pasada cubre el objetivo de `Ronda.md` («detect errors and potential runtime failures») y se ejecuta **por tarea**, no por plan.

### 3.1.1 Pasada determinista (obligatoria)

```powershell
powershell -ExecutionPolicy Bypass -File scripts/audit-tasks.ps1 -Plan .agents/plans/<plan>.md
# subconjunto por ola/tarea:
powershell -ExecutionPolicy Bypass -File scripts/audit-tasks.ps1 -Plan .agents/plans/<plan>.md -Tasks T3,T4
```

| exit | Significado | Veredicto |
|---|---|---|
| `0` | sin hallazgos `Compilation`/`Runtime`/`Logic` (puede haber `Warning`) | continuar a C2 |
| `3` | ≥1 hallazgo bloqueante | `FIX_PLAN` |
| `4` | entrada inválida (plan sin `files_lock`) | `BLOCKED` |

Evidencia: `_evidence/<stamp>-audit/audit.md` (no se commitea). El script audita los archivos declarados en `files_lock` + `files edited/created` de cada `T<n>` y emite la tabla `| Fase | Task | Severidad | Archivo | Líneas | Detector | Hallazgo |`, ya lista para pegar como `Fase {n}` en `Ronda.md`.

| Detector | Severidad | Qué detecta |
|---|---|---|
| `FS-01` | Compilation | archivo declarado en el plan que no existe en disco (deriva plan/disco) |
| `AR-01` | Compilation | Interop/WinForms/`System.Data`/PdfSharp/`System.Drawing` en `Application`/`Domain` |
| `CA-01`/`CA-02` | Compilation / Warning | C# 12 (collection expressions, primary constructors) en net48 |
| `MIG-01` | Compilation | dos `MigrationV<n>.cs` con el mismo correlativo (`__SchemaVersion` ambiguo) |
| `RT-01`/`RT-02` | Runtime | `.Result` / `.Wait()` (deadlock en hilo STA) |
| `RT-03` | Runtime | `async void` (excepción no observable) |
| `RT-04` | Runtime | `catch { }` vacío (excepción tragada, estado inconsistente) |
| `RT-06` | Runtime | `Application.DoEvents` / `Thread.Sleep` (reentrancy, UI congelada) |
| `LN-01` | Runtime | `MessageBox` en capa limpia |
| `LG-05` | Runtime | `First()`/`Single()` sobre secuencia vacía |
| `LG-01`/`LG-02` | Logic / Warning | `DateTime.Now|Today` (falta `IDateTimeProvider`) |
| `LG-03`/`LG-04` | Logic | `SELECT *`, `INSERT` sin lista de columnas (desincronía posicional) |
| `LG-06` | Logic | `DELETE FROM` sobre tabla de `CrdtTables` (tombstone replicado) |
| `TH-01` | Logic | test `Ignore`/`Skip`/`Inconclusive` (suite debilitada) |
| `TH-02` | Warning | assert tautológico |
| `RT-05`/`RT-07`/`RT-08`/`WR-01` | Warning | `SqliteConnection` sin `Dispose`, restauración COM, `GridReader`, `Console.Write` |

Los `Warning` no bloquean el sellado, pero **se listan** en el acta de cierre y en `Ronda {n}`.

### 3.1.2 Revisión semántica por tarea (manual, no delegable)

Para cada `T<n>` lee el archivo **en disco** (nunca el ledger) y confirma en este orden:

1. **Aplicación real**: la corrección de `description`/`validation` está en las líneas que el plan declara, con la lógica exacta (guards, mensajes, ramas de error).
2. **Columnas y esquema**: toda columna/tabla citada en SQL existe en el esquema final (`**/Migrations/` + `__SchemaVersion`); ninguna fue renombrada o eliminada.
3. **Orden de operaciones**: no se asigna un campo/DTO estático antes de terminar de leer el `GridReader`; el commit va después de la última escritura.
4. **Atomicidad**: transacción en `using` con rollback; qué queda inconsistente si la excepción ocurre a mitad (caché, `config`, vault origen vs destino).
5. **Idempotencia**: repetir la operación (upsert, `EnsureUsuarioEnVault`, migración) no duplica filas ni versiones.
6. **STA/COM**: ningún `await`/`.Result` nuevo en el hilo de Excel; estado COM (`ScreenUpdating`, `DisplayAlerts`, `EnableEvents`) restaurado en `finally`.
7. **Nulabilidad**: `First/Single`, `Convert.ToInt32`, `int.Parse`, `(string)range.Value2` sobre datos de Excel vacíos o celdas `null`.
8. **Efectos colaterales de UI**: mensajes obsoletos, combos `Enabled = false`, formularios abiertos con datos viejos.
9. **Higiene de tests**: el diff de `ConversorPDF.Tests` no relaja asserts ni reduce casos; el test nuevo falla por la razón prevista.
10. **Desincronía plan/ejecución**: `files edited/created` ⊆ `files_lock` y coincide con `git show --name-only` del commit `T<n>`.

### 3.1.3 Formato del hallazgo (compatible con `Ronda.md`)

```markdown
### Fase {n} (T{k}): {título del defecto}
- **Ubicación**: `<archivo>:<líneas>`
- **Problema**:
- **Causa raíz**:
- **Severidad**: Compilation | Runtime | Logic | Warning
- **Corrección**: ```csharp / ```sql
- **Origen**: audit-tasks (`RT-01`) | revisión manual §3.1.2 punto {i}
- **Riesgo de rotura**:
```

`{n}` es siempre el número **global** de [`REGISTRY.md`](../../plans/REGISTRY.md); la `T{k}` va entre paréntesis para trazabilidad.

---

## 4. C2 — Auditoría de leases

```powershell
foreach ($c in (git rev-list "$base..HEAD")) {
  git log -1 --pretty=format:'%h %s' $c; git show --name-only --pretty=format: $c; ''
}
```

Regla: para cada commit `T<n>: ...`, los archivos tocados deben ser subconjunto de `files_lock(T<n>)` ∪ `{.agents/plans/<plan>.md}`. Cualquier archivo adicional ⇒ `LEASE-VIOLATION` ⇒ `FIX_PLAN`.

## 5. C2.5 — Validación de numeración y cierre de la fila del registro

Antes de sellar, comprueba que la numeración global es coherente y única:

```powershell
Get-Content .agents\plans\REGISTRY.md | Select-String '^\| `(ronda|fase)`'
Get-Content .agents\plans\<plan>.md | Select-String 'fase:|ronda_base|fase_base'
Get-Content .agents\plans\<plan>.md | Select-String '^\| \d+ \|'      # tabla de olas
```

| Comprobación | Fallo |
|---|---|
| Existe una fila del plan en `REGISTRY.md` §3 | `NUMBERING_MISSING` ⇒ `BLOCKED` |
| `- **fase**:` de cada `T<n>` cae dentro del rango reservado y no se repite | `NUMBERING-CONFLICT` ⇒ `FIX_PLAN` |
| `ronda_base` + olas realizadas ≤ siguiente `ronda` libre | `NUMBERING-CONFLICT` |
| Nº de olas reales ≠ reservadas (por `LEASE-SERIALIZED`) | ajustar al alza en la fila; anotar la desviación en el acta |

Acciones al sellar (sólo Fase C):

1. `reserved|executing` → `closed` en la fila del plan.
2. Escribir `verify`, `HEAD` y el **rango real** de olas/tareas.
3. Actualizar §2 (`ronda`, `fase`) con el último número **efectivamente usado**, no el reservado.
4. Nunca reutilizar ni reciclar números de un plan `abandoned`.

---

## 6. C3 — Invariantes de arquitectura

```powershell
# Capas limpias: Application + Domain de cada slice (mismo criterio que AR-01)
Get-ChildItem -Path ConversorPDF.XLL\Academic -Recurse -Filter *.cs |
  Where-Object { $_.FullName -match '\\(Application|Domain)\\' } |
  Select-String -Pattern 'Excel\.|Microsoft\.Office|DataGridView|System\.Windows\.Forms|System\.Data'

# Riesgos de runtime en el slice
Get-ChildItem -Path ConversorPDF.XLL\Academic -Recurse -Filter *.cs |
  Select-String -Pattern 'new \[\]|DateTime\.Now|\.Result\b|\.Wait\(\)|async void|catch\s*\{\s*\}'
```

Esperado: **0 resultados nuevos respecto a `base`**. Interop/WinForms/`System.Data` en capas limpias, o `DateTime.Now` / `.Result` / `.Wait()` / `catch {}` introducidos por el plan ⇒ `FIX_PLAN`.

`scripts/audit-tasks.ps1` (C1.5) aplica estos mismos patrones acotados a los archivos de cada tarea, así que **la autoridad sobre el código del plan es C1.5**. La deuda preexistente fuera del `files_lock` no bloquea el sellado: se reporta como observación en el acta (hoy: 5 `using System.Data;` en `Academic/Application/Services/**`; las menciones de `DataGridView` en comentarios XML no cuentan).

## 7. C4 — Publicación y documentación

```powershell
git log --oneline origin/refactor/<slug>..HEAD      # los commits DEBEN seguir sin publicar
```

Artefactos exigidos por `development-process` (Finalization):

- `docs/articles/changes/<YYYY-MM-DD>-<slug>.md` (Purpose · Changes · How to use · Test coverage)
- `walkthrough.md`

## 8. C5 — Sellado

1. En la cabecera, `**Estado**: executing` → `closed` (único cambio sobre el ledger de Fase B).
2. Añadir el bloque `## Cierre (Fase C)` con la plantilla literal de [`plans/README.md`](../../plans/README.md) §5.3: `verify Full`, `HEAD`, evidencia, leases, arquitectura, publicación, docs y desviaciones. Añade dos líneas obligatorias:
   - `- **Auditoría de código**: audit-tasks exit <0|3> · <n> archivos · <n> hallazgos (<w> warnings) · _evidence/<stamp>-audit/audit.md`
   - `- **Numeración**: Ronda <r0>–<r1> (Wave 1..<w>) · Fase <f0>–<f1> (T1..<n>) · REGISTRY.md fila <R-xxx>`
3. Registrar el veredicto y la lista de commits auditados.
4. Cerrar la fila de [`REGISTRY.md`](../../plans/REGISTRY.md) (§C2.5).

Ciclo de vida del estado: `draft` → `approved` → `executing` → **`closed`**.

## 9. C6 — Commit local (sin push)

```powershell
git add -- .agents/plans/<plan>.md .agents/plans/REGISTRY.md walkthrough.md docs/articles/changes/<fecha>-<slug>.md
git diff --cached --name-only        # auditar: solo esos archivos
git commit -m "docs(fase-c): cierre <slug> — verify Full exit 0, 0 regresiones"
```

`REGISTRY.md` entra en el commit de cierre porque es el historial de numeración; `_evidence/**` no.

## 10. Contrato de salida

Responde con:

- **Veredicto**: `SELLADO` | `FIX_PLAN` | `BLOCKED`
- **Evidencia**: exit code de `verify.ps1`, `TOTAL` medido vs baseline, `HEAD`, archivos commiteados
- **Auditoría de código**: exit de `audit-tasks.ps1`, nº de archivos y de hallazgos por tarea
- **Hallazgos**: numerados como `Fase {n} (T{k})`, severidad `Compilation|Runtime|Logic|Warning`, ubicación `<archivo>:<líneas>` y riesgo de rotura
- **Tabla resumen** (`Ronda.md`):

  | Fase | Task | Severidad | Archivo | Líneas |
  |---|---|---|---|---|

- **Numeración**: `Ronda <rango>` / `Fase <rango>` y estado de la fila en `REGISTRY.md`
- **Si procede**: ruta del `<YYYY-MM-DD>-<slug>-fix-plan.md` creado

## 11. Prohibiciones

- **No reabrir** el plan cuando hay hallazgos: dos escritores sobre el ledger. Se crea `<YYYY-MM-DD>-<slug>-fix-plan.md` (precedente: `2026-10-27-usuarios-duplicados-subunidad-plan.md` §12).
- **No push, merge ni publish.**
- **No aceptar el ledger** sin re-ejecutar `verify.ps1 -Mode Full` **y** `audit-tasks.ps1` en disco.
- **No aceptar** «el revisor ya lo miró»: sin `audit-tasks.ps1` ejecutado no hay veredicto (`BLOCKED`).
- **No commitear** `_evidence/**` ni archivos fuera del acta de cierre.
- **No sellar** con exit `1`, `2`, `3` ni `4` de `verify.ps1`, ni con exit `3`/`4` de `audit-tasks.ps1`.
- **No inventar** números de `Ronda`/`Fase`: si el plan no está en `REGISTRY.md`, se reporta con `T{k}` y se abre incidencia de numeración.

