# Contrato de Plan para Swarm Multiagente (ConversorPDF)

Este documento especifica el formato obligatorio que todo plan generado por `swarm-planner` (o por Cline) debe cumplir para ser ejecutable por `swarm-parallel-task`.

---

## 1. Reglas Generales de Almacenamiento
- **Ubicación obligatoria**: Todo plan DEBE guardarse en `.agents/plans/<YYYY-MM-DD>-<slug>-plan.md`.
- **Prohibido**: Jamás escribir archivos de plan en la raíz del repositorio (`/`).

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

---

## 4. Estructura de Olas (Waves) y Grupos Paralelos

El plan debe incluir la tabla de olas:
```markdown
## Parallel execution groups
| Wave | Tasks | Can start when |
|---|---|---|
| 1 | T1, T2 | Inmediato (leases disjuntos) |
| 2 | T3 (L0) | T1 y T2 completados |
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
