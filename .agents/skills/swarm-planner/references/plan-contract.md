# Contrato de Plan para Swarm Multiagente (ConversorPDF)

Este documento especifica el formato obligatorio que todo plan generado por `swarm-planner` (o por Cline) debe cumplir para ser ejecutable por `swarm-parallel-task`.

---

## 1. Reglas Generales de Almacenamiento
- **Ubicación obligatoria**: Todo plan DEBE guardarse en `.agents/plans/<YYYY-MM-DD>-<slug>-plan.md`.
- **Prohibido**: Jamás escribir archivos de plan en la raíz del repositorio (`/`).
- **Numeración global**: la reserva de `Ronda` (=`Wave`) y `Fase` (=`T<n>`) DEBE registrarse en `.agents/plans/REGISTRY.md` (ver §7). Sin esa fila el plan no puede sellarse.

---

## 2. Bloque Execution Envelope (Obligatorio al inicio del plan)

Todo plan debe comenzar con el bloque YAML/Markdown de envelope de ejecución:

```markdown
# Plan: <Título descriptivo del feature o refactor>
**Generado**: <YYYY-MM-DD> · **Autor**: Cline/Swarm-Planner · **Estado**: draft | approved | executing | closed

## Execution envelope
- branch: refactor/<slug>
- workspace: <ruta de la copia local VERIFICADA, fuera de OneDrive>
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <workspace> -Plan <plan>
- workspace_guard:
  - min_xll_cs: 300        # el arbol real tiene ~450; un clon de HEAD antiguo tiene ~98
  - min_tests_cs: 100      # el arbol real tiene ~140; un clon antiguo tiene ~15
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: <N medido en el workspace> — NUNCA un numero inventado)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado (un unico build/tests a la vez por workspace; Excel-DNA bloquea bin/)
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
  - waves: <W>               # Ronda <ronda_base>..<ronda_base+W-1>
  - tasks: <N>               # Fase <fase_base>..<fase_base+N-1>
```

---

## 3. Especificación Atómica de Tareas (11 campos obligatorios)

Cada tarea (T1..Tn) debe declarar explícitamente:
1. `id`: Identificador único (ej. `T1`).
2. `depends_on`: Lista de IDs de los cuales depende directamente (`[]` si es raíz).
3. `location`: Carpeta o archivo principal de trabajo.
4. `files_lock`: Lista exacta de archivos a modificar (deben ser disjuntos entre tareas de la misma ola).
5. `description`: Propósito claro y acotado del cambio.
6. `validation`: Criterio falsable y observable de éxito.
7. `verification`: Comando oficial ejecutable (ej. `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "..."`).
8. `evidence`: Ruta del directorio de evidencia generado (`_evidence/<timestamp>/summary.md`).
9. `status`: `pending | executing | red | green | blocked | done`.
10. `log`: Registro conciso de acciones tomadas y salida del comando.
11. `files edited/created`: Lista final de archivos tocados (verificada contra `files_lock`).

### Campo Adicional: Ruling (Solo en caso de disputa o fix)
Si un revisor (`task-reviewer`) rechaza una tarea o surge una ambigüedad, el orquestador tiene un tope de **3 rondas de fix**. Si persiste la diferencia, el orquestador emite un `Ruling` formal:
`Ruling: <decisión tomada> — <justificación técnica> — <costo estimado si la decisión es errónea>`.

### Campo Adicional: `fase` (numeración global obligatoria con registro)
Si el plan tiene fila en `REGISTRY.md`, cada tarea declara su número global de `Fase` justo después de `id`:

```markdown
### T3: <Nombre>
- **id**: T3
- **fase**: 1422
- **depends_on**: [T1, T2]
```

`fase` = `Fase {n}` de `.clinerules/workflows/Ronda.md`. Es el número con el que se reportan los hallazgos (`Fase 1422 (T3)`) y con el que se numera el plan de fix. Debe caer dentro del rango `fase_base..fase_base+tasks-1` reservado en el envelope y no repetirse.

---

## 4. Estructura de Olas (Waves) y Grupos Paralelos

El plan debe incluir la tabla de olas (una `Wave` = una `Ronda` de `Ronda.md`):
```markdown
## Parallel execution groups
| Wave | Ronda | Tasks | Can start when |
|---|---|---|---|
| 1 | 259 | T1, T2 | Inmediato (leases disjuntos) |
| 2 | 260 | T3 (L0) | T1 y T2 completados |
```
Regla de oro: Si dos tareas comparten prefijo en `files_lock`, deben colocarse en olas diferentes o marcarse como serializadas.

---

## 5. Reglas Duras del Gate (aprendidas de la auditoria 2026-10-27)

1. **El baseline nunca se escribe a mano con numeros estimados.** Se genera midiendo el workspace real; `TOTAL` debe reflejar la suite completa.
2. **Suite reducida = fallo.** Si `TOTAL` ejecutado < `TOTAL` del baseline, `verify.ps1 -Mode Full` devuelve `3` (`SUITE_SHRUNK`). Un cambio que hace desaparecer tests no es "verde".
3. **Un solo gate a la vez.** Dos builds concurrentes en el mismo workspace chocan en `bin/` (Excel-DNA `DNA-577853131`). `build-verifier` se invoca serializado.
4. **La evidencia no se commitea.** `_evidence/` esta en `.gitignore`; se enlaza por ruta en el ledger y se commitean solo el plan y los archivos del `files_lock`.
5. **Workspace verificado, no asumido.** `scripts/preflight.ps1` debe devolver `0` antes de la primera ola. Verifica tamano del arbol (`min_xll_cs`/`min_tests_cs`), presencia del slice y coherencia baseline/arbol.
6. **El workspace debe provenir del arbol real.** Un `git clone` solo trae HEAD: si el working tree tiene cambios sin commitear, el clon es un snapshot incompleto y todo el baseline sera invalido.

---

## 7. Registro de numeración (Ronda ⇔ Wave · Fase ⇔ Task)

Los planes swarm y `.clinerules/workflows/Ronda.md` describen el mismo objeto con dos vocabularios:

| `Ronda.md` / `implementation_plan.md` / `walkthrough.md` | Plan swarm | Nivel |
|---|---|---|
| `Ronda {n}: {título}` | `Wave {n}` (parallel execution group) | Grupo paralelo |
| `Fase {n}: {título}` | `T{n}: {Nombre}` | Tarea atómica |

La numeración es **global, monótona y append-only**; su fuente única es `.agents/plans/REGISTRY.md`:

1. **Reserva (Fase A, `swarm-planner`)**: leer `ronda`/`fase` de §2, reservar `ronda+1..ronda+W` y `fase+1..fase+N`, declararlos en el bloque `numbering:` del envelope, añadir `- **fase**: <n>` a cada tarea y una fila `reserved` en §3.
2. **Ejecución (Fase B, `swarm-parallel-task`)**: no renumera. Si serializa olas por colisión de leases, lo registra (`LEASE-SERIALIZED`) y el cierre ajusta el rango al alza.
3. **Cierre (Fase C, `swarm-verify`)**: valida que cada `fase` cae en el rango reservado y no se repite, cierra la fila (`closed`), actualiza contadores con el último número **usado** y lo refleja en el bloque `## Cierre (Fase C)`.
4. **Plan de fix**: rango nuevo + referencia a las `Fase {n}` corregidas. Un plan `abandoned` no libera sus números.
5. **Prohibido**: inventar contadores, reciclar números o numerar hallazgos sin fila previa en el registro.
