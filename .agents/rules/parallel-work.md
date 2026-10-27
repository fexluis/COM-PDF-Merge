# Reglas de Trabajo Paralelo y Sistema de Leases (Modelo Superpowers)

> [!IMPORTANT]
> **Trigger**: `always_on`. Aplica a cualquier plan, tarea o subagente que opere sobre el repositorio ConversorPDF.

---

## 1. Roles y Permisos de Escritura

| Rol | Permisos de Escritura | Responsabilidad Principal |
|---|---|---|
| `plan-reviewer` | ❌ Solo Lectura | Detectar dependencias faltantes, orden no ejecutable, criterios no verificables o solapes de lease en el plan antes de iniciar ejecución. |
| `tdd-red-writer` | Únicamente `ConversorPDF.Tests/**` | Escribir tests unitarios que fallen por la razón exacta esperada (fase RED) y pegar la salida TRX/summary en la evidencia. |
| `implementer` | Únicamente su `files_lock` asignado | Implementar el código mínimo de producción para pasar a GREEN sin debilitar tests y sin tocar archivos L0. |
| `integrator` | Archivos de **Lease L0** exclusivamente | Crear esqueletos de tipos nuevos (`integrator-first`), actualizar `<Compile Include>` en el csproj, registrar servicios, asignar migraciones de base de datos. |
| `task-reviewer` | ❌ Solo Lectura | Revisar el diff contra especificación (spec compliance) e higiene/calidad del código (YAGNI, invariantes de arquitectura). |
| `build-verifier` | Únicamente `_evidence/**` | Ejecutar `scripts/verify.ps1`, generar TRX y `summary.md`, certificar transiciones RED/GREEN y validar regresiones vs baseline. |

---

## 2. Definición Estricta de Leases

### Lease L0 — Exclusivo Global (Solo `integrator`, 1 tarea a la vez)
Los siguientes archivos y patrones son considerados caminos calientes globales. Ningún `implementer` ni `tdd-red-writer` puede tenerlos en su `files_lock`:
- `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- `ConversorPDF.XLL/Core/Helpers/App.cs`
- `ConversorPDF.XLL/**/Migrations/MigrationV*.cs`
- `ConversorPDF.XLL/**/Persistence/SchemaMigrator.cs`
- `ConversorPDF.XLL/Ribbon/RibbonController.cs`
- Composition Root / Registro central de IoC/DI
- `docs/toc.yml`

### Lease L1 — Por Archivo o por Slice
- Archivos dentro de `ConversorPDF.XLL/Academic/<Slice>/**` o carpetas de dominio.
- Cuando dos tareas operan dentro de la misma carpeta de un slice, el lease DEBE ser granular a nivel de archivo físico (`files_lock: [Academic/Genetica/Application/Foo.cs]`), nunca a nivel de carpeta entera.

### Lease L2 — Solo Lectura
- `.agents/**` (excepto el orquestador que actualiza el plan/ledger), handoff docs, y `docs/**` (excepto `docs/articles/changes/` reservado al cierre por Cline).

---

## 3. Concurrencia y Solapes de Lease
1. **Límite de Escritores**: `max_parallel_writers: 2`. Los agentes de solo lectura (`plan-reviewer`, `task-reviewer`) pueden ejecutarse concurrentemente sin límite.
2. **Detección de Solape**: Si dos tareas en la misma ola contienen rutas idénticas o prefijos de ruta que se intersequen en `files_lock`, el orquestador tiene terminantemente prohibido ejecutarlas en paralelo. Debe:
   - Serializar la ejecución (Ola posterior).
   - Registrar explícitamente en el log: `LEASE-SERIALIZED T<m> <- T<n> (motivo: colisión en <ruta>)`.

---

## 4. Reglas Críticas de Arquitectura y Lenguaje

1. **C# en Fase RED (El RED debe compilar)**:
   - Dado que el compilador de C# requiere que los tipos invocados existan para compilar los tests, si una tarea introduce nuevas clases o interfaces, el `integrator` debe crear primero el esqueleto mínimo con métodos que lancen `NotImplementedException` y agregarlo al `.csproj` (`integrator-first`), ANTES de que el `tdd-red-writer` escriba los tests.
2. **Asignación de Migraciones de Base de Datos**:
   - Nunca predecir o asignar versiones de migración `V<n>` en paralelo.
   - El `integrator` es el único facultado para inspeccionar las migraciones existentes en `Migrations/` y `__SchemaVersion`, asignando el correlativo exacto en el momento atómico de escribir.
3. **Control de Versión de la Aplicación**:
   - `APP_VERSION` en `App.cs` y `<Version>` en el `.csproj` solo son editados por el `integrator` durante el cierre de una release oficial, nunca durante el desarrollo de una feature.
4. **Propiedad del Plan y Ledger**:
   - El archivo del plan en `.agents/plans/` tiene un único escritor activo por fase:
     - Durante la Fase A: Cline.
     - Durante la Fase B: El orquestador de Antigravity (`swarm-parallel-task`).
     - Durante la Fase C: Cline (lectura y cierre final).
   - Ningún subagente implementador escribe directamente en el archivo del plan.

---

## 5. Workspace, Gate Único y Evidencia

1. **Workspace verificado, no asumido**:
   - Antes de la primera ola: `powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ruta> -Plan <plan>` debe devolver `0`.
   - El workspace debe ser una copia del **árbol de trabajo real** (`min_xll_cs` / `min_tests_cs` del envelope). Un `git clone` de un repo con cambios sin commitear es un snapshot incompleto y se rechaza (`PREFLIGHT_FAILED`).
   - Se permite OneDrive solo con `-AllowOneDrive` y aviso explícito en el ledger (locks de MSBuild/VSTest).

2. **Gate serializado (un solo `build-verifier` a la vez)**:
   - Dos builds concurrentes en el mismo workspace fallan con `DNA-577853131` (Excel-DNA no puede escribir `bin/x64/Debug/*.xll`).
   - El orquestador nunca lanza dos `verify.ps1` simultáneos, ni siquiera en modo `Build`.

3. **Evidencia**:
   - Vive en `_evidence/<timestamp>/` (ignorada por Git) y se enlaza por ruta en el ledger.
   - **Nunca se commitea**: `git add` se limita a los archivos del `files_lock` + el plan.

4. **Baseline íntegro**:
   - `TOTAL` del baseline debe reflejar la suite completa del workspace.
   - `Full` con `TOTAL` ejecutado menor que el baseline se considera fallo (`SUITE_SHRUNK`), no éxito.
