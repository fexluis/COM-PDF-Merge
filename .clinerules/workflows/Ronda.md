# Workflow: Ronda

## Trigger
Use this workflow when the user asks to review code, detect errors, or validate a recent implementation.

## Objective
Detect errors and potential runtime failures in the requested C#/SQL code, explain the root cause of each issue, and propose concrete corrections with code snippets.

## Equivalencia de numeración (leer antes de empezar)

`Ronda {n}` ≡ `Wave {n}` (parallel execution group) y `Fase {n}` ≡ `Task T{n}` del plan swarm. La numeración es **global, monótona y append-only**; su fuente única es `.agents/plans/REGISTRY.md`.

| Este workflow | Plan swarm | Nivel |
|---|---|---|
| `## Ronda {n}: {título}` | `Wave {n}` (parallel execution group) | Grupo de ejecución paralela |
| `## Fase {n}: {título}` | `### T{n}: {Nombre}` | Tarea atómica |
| `| Fase | Severidad | Archivo | Líneas |` | Tabla de `scripts/audit-tasks.ps1` | Reporte |

Para continuar la numeración lee, en este orden:

1. `.agents/plans/REGISTRY.md` §2 (contadores `ronda` / `fase`) y §3 (fila del plan).
2. La tabla de olas del plan (`| Wave | Ronda | Tasks |`) y el campo `- **fase**:` de cada tarea.
3. La última `Ronda {n}` / `Fase {n}` en `implementation_plan.md` y `walkthrough.md`.

Continúa desde el **mayor** de esos valores. Nunca reutilices ni inventes números; si el plan no tiene fila en el registro, usa `T{k}` como referencia y declara la incidencia de numeración.

## Steps
1. **Identify the scope**
Buscar la última ronda y última fase numerada (ver «Equivalencia de numeración»).
Ronda #: Enumera la ronda para iniciar la revisión de las tareas de la ola anterior; continúa la numeración exacta:

## Ronda {n}: {tittle} (Error/problem title)

Contexto de la ronda anterior:
- Ola revisada: `Wave {n-1}` (Tasks: T{a}..T{b} → Fases {f-a}..{f-b})
- Fases aplicadas y verificadas: {lista}
- Fases retiradas o falso positivo: {lista, si aplica}
- Estado de corrección pendiente: {fases incompletas / fases abiertas}

2. **Detect runtime/logic errors and Propose concrete corrections**
Revisa exhaustivamente el código de las Tasks aplicadas en la ola anterior (`T1: <Nombre> {T-?}...{n}`), que equivalen a las fases `{n-?}...{n}`. Confirma que cada corrección se aplicó como fue especificada (líneas, lógica, validaciones) y no introdujo regresiones.

Pasada determinista (solo si el modo permite ejecutar scripts; no compila ni edita código):

```powershell
powershell -ExecutionPolicy Bypass -File scripts/audit-tasks.ps1 -Plan .agents/plans/<plan>.md -Tasks T1,T2
```

`exit 0` = sin hallazgos bloqueantes · `exit 3` = hay hallazgos `Compilation`/`Runtime`/`Logic` (cada uno se convierte en una `Fase {n}`) · `exit 4` = entrada inválida. El informe queda en `_evidence/<stamp>-audit/audit.md` con la tabla `| Fase | Task | Severidad | Archivo | Líneas |` lista para usar.

Si detectas errores nuevos o incompletos (compilación, runtime, lógica, UX/warning), enuméralos siguiendo el formato de fase, continuando la numeración exacta desde la última fase conocida.

Detecta cualquier error o problema que pueda generar fallos en su ejecución
   - For each error, explain the cause and provide the corrected C# or SQL snippet.
   - Keep changes minimal and non-destructive.

### Checklist de fallos de ejecución (por Task)

- **Bloqueos STA**: `.Result`, `.Wait()`, `Thread.Sleep`, `Application.DoEvents`, `async void` nuevos.
- **Fronteras de capa**: `Excel.` / `Microsoft.Office.Interop` / `DataGridView` / `System.Windows.Forms` / `System.Data` en `Application` o `Domain`.
- **SQL**: columna inexistente tras una migración; `SELECT *`; `INSERT` sin lista de columnas; `DELETE FROM` en tabla de `CrdtTables` (tombstone); `MigrationV<n>` duplicada.
- **Atomicidad**: transacción con rollback; qué queda inconsistente si falla a mitad (caché, `config`, vault origen vs destino).
- **Idempotencia**: repetir upsert/migración no duplica filas ni versiones.
- **Nulabilidad**: `First()`/`Single()`, `Convert.ToInt32`, celdas de Excel vacías.
- **COM**: `ScreenUpdating` / `DisplayAlerts` / `EnableEvents` restaurados en `finally`; RCW liberados.
- **Tests**: sin asserts relajados, sin `Ignore`/`Skip`, sin `--filter` que excluya casos existentes.
- **Plan vs disco**: `files edited/created` ⊆ `files_lock` y coincide con `git show --name-only` del commit `T{k}`.

## Fase {n}: {tittle} (Error/problem title)

**Task:** T{k}: {Nombre} · Wave {w} (nº global `Fase {n}` según `REGISTRY.md`)

**Ubicación:** {archivo}, líneas {n}-{m} (File and line reference)

**Problema:** {descripción} 

**Causa:** {raíz} (Cause explanation)

**Corrección:**
{código} (Corrected code snippet)

---
**Severity: Compilation / Runtime / Logic / Warning**

Al final, tabla resumen:

| Fase | Task | Severidad | Archivo | Líneas |
|------|------|-----------|---------|--------|

Y conclusión breve.


3. **Save AI tokens and reduce costs**
Responde con la mínima cantidad de tokens:
- Sin relleno, sin cortesías
- Sin lenguaje emocional
- Frases cortas y directas
- Prefiere fragmentos en vez de oraciones completas
- Quita palabras innecesarias (artículos, relleno, repetición)
- Enfócate solo en resolver el problema
- Entrega solo información útil

4. **Verification Plan**
Verifica que las fases no introducen regresiones.
Conclusión con decisión final Ronda {n}:
a. Si se requieren correcciones / Nuevas Fases
b. Sin Fases nuevas obligatorias / No se requieren más correcciones / Fase opcional / Recomendar pasar a build release + smoke test.
 
5. **Restricciones**
- No se programan parches ad-hoc ni excepciones para casos específicos (hardcodes de IDs, fechas mágicas o bypasses puntuales). Se debe diagnosticar y resolver siempre la causa raíz a nivel de modelo y arquitectura para que el sistema sea genérico, determinista y el error quede estructuralmente mitigado para siempre.
- Requiere modo Act para ejecutar build/tests/smoke.
- No compilar, No editar archivos. Esperar aprobación.
- `scripts/audit-tasks.ps1` es solo lectura de código (escribe únicamente `_evidence/`); si el modo no permite ejecutarlo, la revisión es 100 % manual y así se declara.

6. **Trazabilidad de la numeración**
- Reporta siempre `Fase {n} (T{k})` con el número global del registro.
- Si la ronda se cierra fuera de un plan swarm (revisión manual), actualiza una sola vez §2 de `.agents/plans/REGISTRY.md` con la última `ronda`/`fase` usada; dentro de un plan lo hace la Fase C (`swarm-verify`).
- No inventes números ni recicles rangos: el registro es append-only.
- Antes de responder, verifica los umbrales de `.agents/rules/history-rotation.md`; si `implementation_plan.md` o `walkthrough.md` los superan, ejecuta la escisión con `.agents/tools/split-history.ps1` y actualiza el índice antes de cerrar la ronda.


