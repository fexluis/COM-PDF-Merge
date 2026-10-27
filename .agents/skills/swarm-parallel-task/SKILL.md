---
name: swarm-parallel-task
description: >
  Only to be triggered by explicit /swarm-parallel-task or /parallel-task commands.
  Orchestrates parallel subagents in dependency waves with strict lease ownership, RED->GREEN evidence, and atomic commits.
---

# Parallel Task Executor (Swarm Parallel Task - Modelo Superpowers)

Eres el **Orquestador Central** de la ejecución del plan en ConversorPDF. Tu función es coordinar la ejecución en olas utilizando los subagentes especializados (`plan-reviewer`, `tdd-red-writer`, `implementer`, `integrator`, `task-reviewer`, `build-verifier`).

> [!CRITICAL]
> **Regla de Ownership**: Eres el **único escritor autorizado** del archivo de plan (`.agents/plans/<plan>.md`) durante la Fase B de ejecución. Escribes los registros en el ledger y ejecutas commits locales atómicos por tarea. **NUNCA hagas `git push`**.

---

## Flujo Operativo en 10 Pasos

### Paso 1: Parseo y Pre-Flight de Validación
1. Leer el archivo de plan indicado en `.agents/plans/<plan>.md`.
2. Verificar la presencia y validez del bloque `## Execution envelope`.
3. Validar pre-requisitos indispensables en disco:
   - Script de verificación: `scripts/verify.ps1` existe.
   - Archivo de baseline: `tests/baseline/test_baseline.txt` existe.
   - Workspace activo: el declarado en el envelope, validado con `powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ruta> -Plan <plan>` (debe devolver `0`). No se asume ninguna ruta: debe ser una copia del arbol real (`min_xll_cs`/`min_tests_cs`).
   - Gate serial: un unico `build-verifier` a la vez (dos builds concurrentes rompen con `DNA-577853131`).
   - Definición de los 6 agentes en `.agents/agents/`.
4. Si falta alguno de estos elementos, abortar inmediatamente con estado `PREFLIGHT_FAILED`.

### Paso 2: Gate Inicial de Auditoría (`plan-reviewer`)
1. Invocar al subagente `plan-reviewer` pasando la ruta completa del plan.
2. Analizar el informe devuelto:
   - Si el dictamen es `BLOCKED` o contiene hallazgos `Critical`, detener la ejecución de inmediato y notificar los problemas de diseño.
   - Si el dictamen es `APROBADO`, continuar al Paso 3.

### Paso 3: Cálculo de Olas y Detección de Solape de Leases
1. Agrupar las tareas desbloqueadas en olas (`Waves`) respetando su grafo `depends_on`.
2. **Auditoría de colisiones de `files_lock`**:
   - Para toda tarea dentro de la misma ola, comparar sus listas de `files_lock`.
   - Si dos tareas comparten una ruta o un prefijo de archivo, **está terminantemente prohibido ejecutarlas en paralelo**.
   - Acción: Mover una de las tareas a una ola posterior y registrar en el log:
     `LEASE-SERIALIZED T<m> <- T<n> (motivo: colisión en <ruta>)`.
3. Respetar el límite global: `max_parallel_writers: 2`.

### Paso 4: Estrategia `integrator-first` (Para tipos nuevos)
Si una tarea de la ola introduce nuevas clases, interfaces o tipos en C#:
1. Invocar primero al `integrator` de forma serial para crear el esqueleto básico con `NotImplementedException` y registrarlo en `ConversorPDF.XLL.csproj`.
2. Verificar con `verify.ps1 -Mode Build` que el proyecto compila con 0 errores antes de continuar a la fase de pruebas.

### Paso 5: Fase RED Paralela (`tdd-red-writer`)
1. Para cada tarea de la ola con lease disjunto, invocar a `tdd-red-writer`.
2. El `tdd-red-writer` escribe los tests en `ConversorPDF.Tests/**`.
3. Invocar al `build-verifier` para certificar el estado RED:
   - Ejecuta `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "..."`.
   - Verifica que el test falle por la razón exacta prevista.
   - Registra la evidencia en el plan: `status: red`.

### Paso 6: Fase GREEN Paralela (`implementer`)
1. Invocar al `implementer` para cada tarea (máximo 2 en paralelo), restringido estrictamente a su `files_lock`.
2. El implementador desarrolla el código de producción sin tocar archivos L0 ni debilitar tests.
3. Invocar al `build-verifier` para certificar el paso a GREEN:
   - Ejecuta `verify.ps1 -Mode Fast -Filter "..."` (debe dar 0 errores).

### Paso 7: Integración Serial L0 (`integrator`)
Si la tarea requirió modificaciones en caminos calientes (migraciones SQLite, registro en DI o ribbon):
1. Invocar al `integrator` de manera aislada y serial.
2. Verificar con `verify.ps1 -Mode Build`.

### Paso 8: Revisión de Dos Etapas (`task-reviewer`)
1. Invocar al `task-reviewer` para auditar la tarea completada:
   - Etapa 1: Spec Compliance (criterios de aceptación satisfechos).
   - Etapa 2: Quality & Architecture (cero Interop/WinForms en capas limpias, cero C# 12, Single Type Per File).
2. **Manejo de Hallazgos**:
   - Si hay hallazgos `Critical`, reabrir ciclo de corrección (máximo 3 rondas).
   - Si tras 3 rondas persiste el desacuerdo, el orquestador emite un `Ruling` formal en el ledger:
     `Ruling: <decisión> — <justificación técnica> — <costo si es erróneo>`.

### Paso 9: Certificación de la Suite, Commit Atómico y Actualización del Ledger
1. Invocar al `build-verifier` con la suite completa:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
   ```
   Confirmar exit code `0` (0 regresiones vs baseline).
2. Actualizar el ledger de la tarea en el plan con:
   - Estado: `done`
   - Enlace de evidencia: `_evidence/<timestamp>/summary.md`
   - Archivos finales editados/creados.
3. Si Git está disponible, realizar el commit local de la tarea **solo** con los archivos del lease y el plan:
   ```powershell
   git add -- <archivos del files_lock> .agents/plans/<plan>.md
   git diff --cached --name-only   # verificar que no entra nada fuera del lease
   git commit -m "T<n>: <resumen de la tarea>"
   ```
   *(La evidencia `_evidence/**` NO se commitea: esta en .gitignore y se enlaza por ruta en el ledger. Recordatorio: NUNCA hacer push).*

### Paso 10: Cierre de Ola y Fallback Local
1. Avanzar a la siguiente ola hasta agotar todas las tareas del plan.
2. Al finalizar todas las olas, ejecutar una verificación Full final.
3. **Fallback sin subagentes**: Si la plataforma no soporta `invoke_subagent`, recurrir automáticamente a la skill `executing-plans-local` manteniendo la misma disciplina de leases, comandos de verificación y ledger.
