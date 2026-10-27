---
name: executing-plans-local
description: >
  Fallback sequential plan execution engine when invoke_subagent is unavailable.
  Executes tasks in topological order following the same lease rules, verification commands, and ledger tracking.
---

# Local Plan Execution Engine (Fallback sin Subagentes)

Esta habilidad se activa cuando el entorno no cuenta con soporte nativo para invocar subagentes paralelos (`invoke_subagent`). Permite ejecutar el mismo plan de desarrollo (`.agents/plans/<plan>.md`) de forma secuencial y determinista, manteniendo la misma disciplina estricta de leases, verificación RED->GREEN y registro en el ledger.

---

## Procedimiento de Ejecución Secuencial

### 1. Pre-Flight
- Ejecutar `powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <workspace> -Plan <plan>` y exigir exit `0`. No asumir ninguna ruta: el workspace debe ser una copia del arbol real (`min_xll_cs`/`min_tests_cs`).
- Confirmar existencia de `scripts/verify.ps1` y `tests/baseline/test_baseline.txt`.
- Verificar que el repositorio de Git tenga el working tree limpio.

### 2. Orden Topológico de Tareas
- Leer el plan y construir la cola de ejecución resolviendo el grafo de dependencias `depends_on`.
- Ejecutar las tareas una por una de forma estrictamente secuencial.

### 3. Ciclo TDD por Tarea
Para cada tarea `T<n>`:
1. **Paso Esqueleto (si aplica)**: Si la tarea introduce tipos nuevos, crear el archivo físico con métodos que arrojen `NotImplementedException` e incluirlo en `ConversorPDF.XLL.csproj`. Compilar con `scripts/verify.ps1 -Mode Build`.
2. **Paso RED**:
   - Escribir los tests unitarios en `ConversorPDF.Tests/**`.
   - Ejecutar: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~<ClaseBajoPrueba>"`.
   - Confirmar el fallo por la razón exacta prevista y registrar evidencia.
3. **Paso GREEN**:
   - Implementar el código mínimo en producción dentro del `files_lock` asignado.
   - Ejecutar: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~<ClaseBajoPrueba>"`.
   - Confirmar que los tests pasen a verde.
4. **Paso Verificación de Regresiones**:
   - Ejecutar: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`.
   - Confirmar exit code `0` (0 regresiones vs baseline).
5. **Paso Ledger y Commit Local**:
   - Actualizar el estado de la tarea en el plan a `status: Completed`.
   - Vincular la ruta de `_evidence/<timestamp>/summary.md`.
   - Ejecutar commit local: `git commit -m "T<n>: <resumen>"`. (Sin push).
