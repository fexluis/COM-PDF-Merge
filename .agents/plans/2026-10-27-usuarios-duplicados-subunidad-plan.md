# Plan: Usuarios duplicados y cambio de subunidad que no persiste

**Generado**: 2026-10-27 · **Autor**: Cline/Swarm-Planner · **Estado**: draft
**Workspace verificado**: `C:\dev\ConversorPDF` (clon fiel de `d32115e`, 452/138 `.cs`, baseline `TOTAL: 939`)

## Execution envelope
- branch: refactor/academic-modularization
- workspace: C:\dev\ConversorPDF
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <workspace> -Plan .agents/plans/2026-10-27-usuarios-duplicados-subunidad-plan.md
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: 939 · medido 2026-10-27 sobre d32115e)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado (un unico build-verifier a la vez)
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0 (ampliado para este plan):
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Core/Helpers/App.cs
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV*.cs
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs
  - ConversorPDF.XLL/Ribbon/RibbonController.cs
  - ConversorPDF.XLL/Core/Sync/SyncInitializer.cs
  - ConversorPDF.XLL/Core/Sync/SubunidadResolver.cs
  - Composition Root / IoC Container
  - docs/toc.yml
- stop_conditions:
  - Accion destructiva sobre el vault cifrado sin respaldo previo
  - Intento de push, merge o publicacion a remoto
  - Reescritura de IDs en una tabla CRDT sin Ruling documentado y backup del vault
  - Plan irrecuperable tras 3 rondas de fix
- ledger_owner: swarm-parallel-task
- no_push: true

---

## 1. Sintoma reportado

1. En **Gestion de Usuarios** aparecen filas repetidas del mismo `NombreUsuario` (p. ej. `Usuario`/`Usuario Predeterminado` dos veces; `lucs bechara`/`Luis` dos veces).
2. Al cambiar **Subunidad** de un usuario y pulsar **Guardar**, el cambio se ve en pantalla, pero tras **cerrar Excel y volver a abrir** no persiste.

## 2. Diagnostico (evidencia en codigo, no suposiciones)

| # | Defecto | Evidencia |
|---|---|---|
| **D1** | **El guardado no persiste el vault cifrado.** `FormUsuarios` escribe `config` con SQL directo sobre la conexion temporal y nunca re-encripta el `.db.enc` | `FormUsuarios.cs:462` (`db.Execute("UPDATE config SET subunidad = @Subunidad", ...)`); `AcademicRepository.cs:28` ("los metodos de escritura requieren llamar a `ReEncryptAndPersist()` explicitamente"); `AcademicDbSession` no expone `Persist()` (solo `MarkChangesPersisted`/`NotifyChangesCommitted` → `SyncContext.Scheduler?.MarkDirty()`, `AcademicDbSession.cs:69-78`); unico punto que persiste hoy: `SqliteHelper.UpdateVaultDirectly` (`SqliteHelper.cs` ~962), usado solo en `InitializeDatabase` |
| **D2** | **Un vault por subunidad**: cambiar de subunidad cambia de archivo de base de datos. La fila editada queda en el vault origen; al reiniciar se abre el vault destino | `SqliteHelper.GetDbPathForSubunidad()` → `AcademicData_<SUBUNIDAD>.db.enc`; `SubunidadResolver.cs:44-49` resuelve la subunidad desde `usuarios`; `AcademicDbSession` usa `SqliteHelper.GetContextForPath()` (un contexto por vault) |
| **D3** | **Duplicados por identidad no determinista**: `usuarios` no tiene UNIQUE por `nombreusuario` (se elimina a proposito para CRDT) y cada insercion genera GUID nuevo; el sync multi-dispositivo inserta la misma cuenta con otro `id`. Ademas el resolver elige fila con `LIMIT 1` sin `ORDER BY` | `SyncInitializer.cs:287-291` (`DROP INDEX ... idx_usuarios_nombreusuario`), `:558` (se elimino la deduplicacion `MIN(id)`); `AcademicRepository.cs:1383` (`GetUsuarios` sin dedupe), `:1420` (`Guid.NewGuid()`); `MigrationV41_DeduplicateUsuarios` (limpieza **one-shot**); `AcademicSchema.sql:252-264` (indice no unico); seeds `usuarios.sql` usan `WHERE NOT EXISTS (... nombreusuario = ...)` (no son la causa) |

**Correlacion entre ambos sintomas:** si existen dos filas con el mismo `nombreusuario`, guardar actualiza **una** (`WHERE id = @Id`, `AcademicRepository.cs:1440-1442`) y `SubunidadResolver` puede leer **la otra** al arrancar ⇒ el cambio "desaparece". D1 impide ademas que `config.subunidad` sobreviva al cierre. Por eso el plan ataca D1 + D2 + D3.

## 3. Decisiones de diseno

1. **Persistir desde el contexto, no con SQL suelto.** El form recibira la sesion (`AcademicDbSession`) y llamara `Persist()` tras guardar; el SQL directo de `config` se sustituye por `IConfigRepository.UpdateSubunidad`.
2. **Identidad determinista de usuario** (`id = "usr_" + lower(nombreusuario)`). Verificado: **no hay** `REFERENCES usuarios`, asi que reescribir PKs no rompe integridad. Mismo `nombreusuario` en dos PCs ⇒ mismo PK ⇒ CR-SQLite converge en vez de duplicar.
3. **Dedupe en lectura + dedupe recurrente.** El grid nunca muestra repetidos (`GetUsuarios` con `ROW_NUMBER`) y la limpieza deja de ser one-shot (se ejecuta tras el merge de sync).
4. **Resolver determinista** (`ORDER BY activo DESC, updated_at DESC, id`) para que con datos historicos el resultado sea estable.
5. **Mover usuario entre vaults** solo si T2 demuestra que el cambio del usuario propio debe viajar al vault destino. Si no se demuestra, T6 se cierra con `Ruling` y sale del alcance.
6. **Reescritura de IDs = operacion delicada** (tabla CRDT): exige backup del vault, `PRAGMA integrity_check` posterior y tarea `integrator` aislada.

## 4. Grafo de dependencias

```
T1 (integrator: contratos + wiring RibbonController + esqueletos + csproj)
 ├── T2 (RED: persistencia E2E + duplicados + resolver)
 │     └── T3 (implementer: Persist() + ConfigRepository + FormUsuarios sin SQL directo)
 │           └── T6 (implementer: mover usuario entre vaults)
 ├── T4 (implementer: identidad determinista + GetUsuarios dedupe + seeds)
 │     ├── T7 (integrator: MigrationV79 normalizar IDs)
 │     │     └── T8 (integrator: dedupe idempotente post-sync)
 │     └── T5 (integrator: SubunidadResolver determinista)
T9 (gate final + review)                                   [deps: T1..T8]
```

---

## 5. Tareas

### T1: Contratos, wiring del formulario y esqueletos (integrator-first, L0)
- **depends_on**: []
- **location**: `ConversorPDF.XLL/Academic/Application/{Interfaces,Services}`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/{Context,Repositories}`, `ConversorPDF.XLL/Ribbon/RibbonController.cs`, `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`
- **files_lock**:
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/Ribbon/RibbonController.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Context/AcademicDbSession.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/ConfigRepository.cs`
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IConfigRepository.cs`
  - `ConversorPDF.XLL/Academic/Application/Interfaces/IMoverUsuarioSubunidadService.cs` (nuevo)
  - `ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs` (nuevo, lanza `NotImplementedException`)
- **description**: (a) `AcademicDbSession.Persist()` que lance `NotImplementedException`; (b) `IConfigRepository.UpdateSubunidad(string)` + stub en `ConfigRepository`; (c) `IMoverUsuarioSubunidadService.Mover(string usuarioId, string subunidadDestino)` + clase stub; (d) cambiar `FormUsuarios` para recibir `AcademicDbSession` en vez de `Func<IDbConnection>` y actualizar el único call-site `RibbonController.cs:1256`; (e) registrar los 2 archivos nuevos en el `.csproj` sin duplicar (evitar CS2002). Un tipo por archivo.
- **validation**: `verify.ps1 -Mode Build` exit 0.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build`
- **evidence**: `_evidence/20261027-125347`
- **status**: Completed
- **log**: Integrator-first completado. Creados IMoverUsuarioSubunidadService y MoverUsuarioSubunidadService. Agregado Persist() stub a AcademicDbSession y UpdateSubunidad a IConfigRepository/ConfigRepository. Actualizado constructor de FormUsuarios y call-site en RibbonController. Registrados tipos en csproj. Build OK (0 errores).
- **files edited/created**:
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.XLL/Ribbon/RibbonController.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Context/AcademicDbSession.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/ConfigRepository.cs`
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IConfigRepository.cs`
  - `ConversorPDF.XLL/Academic/Application/Interfaces/IMoverUsuarioSubunidadService.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs`

### T2: RED — persistencia, duplicados y resolver (tdd-red-writer)
- **depends_on**: [T1]
- **location**: `ConversorPDF.Tests/Integration/Usuarios/`, `ConversorPDF.Tests/Unit/Usuarios/`
- **files_lock**:
  - `ConversorPDF.Tests/Integration/Usuarios/UsuariosPersistenciaTests.cs`
  - `ConversorPDF.Tests/Integration/Usuarios/UsuariosDuplicadosTests.cs`
  - `ConversorPDF.Tests/Unit/Usuarios/SubunidadResolverTests.cs`
- **description**:
  - **(a) Persistencia E2E**: crear contexto sobre vault temporal (`EncryptedSQLiteContext` + `AcademicDbSession` interno), insertar usuario, cambiar `subunidad`, llamar `Persist()`, disponer, **reabrir** el mismo vault y assert: `usuarios.subunidad` y `config.subunidad` tienen el valor nuevo. Reproduce el "cerrar y abrir Excel".
  - **(b) Duplicados**: insertar dos filas con el mismo `nombreusuario` e `id` distinto (simula merge CRDT) ⇒ `GetUsuarios()` devuelve **una** fila y es la que tiene firma/perfil/`updated_at` más reciente.
  - **(c) Resolver**: con duplicados, `SubunidadResolver` elige activo + `updated_at` más reciente (orden determinista).
  - xUnit + FluentAssertions 6.12; nombres `{Accion}_When_{Condicion}_Should_{Resultado}`.
- **validation**: compilan y fallan por la razón esperada (`NotImplementedException` o aserción), nunca por error de compilación.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~UsuariosPersistenciaTests|FullyQualifiedName~UsuariosDuplicadosTests|FullyQualifiedName~SubunidadResolverTests"`
- **evidence**: `_evidence/20261027-125730`
- **status**: Completed
- **log**: Fase RED certificada. 3 archivos de prueba creados. Los 4 tests compilaron correctamente y fallaron exactamente por las razones previstas: NotImplementedException en Persist(), HaveCount(1) falló con 2 filas por falta de deduplicación, result.Value falló al esperar id determinista 'usr_pedroperez' recibiendo GUID aleatorio, y SubunidadResolver retornó ANTIOQUIA en vez de VALLE por ausencia de ORDER BY.
- **files edited/created**:
  - `ConversorPDF.Tests/Integration/Usuarios/UsuariosPersistenciaTests.cs`
  - `ConversorPDF.Tests/Integration/Usuarios/UsuariosDuplicadosTests.cs`
  - `ConversorPDF.Tests/Unit/Usuarios/SubunidadResolverTests.cs`

### T3: Persistir el vault y eliminar el SQL directo del formulario (implementer)
- **depends_on**: [T2]
- **location**: `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Context/AcademicDbSession.cs`, `.../Repositories/ConfigRepository.cs`, `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Context/AcademicDbSession.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/ConfigRepository.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`
- **description**: implementar `Persist()` = `PRAGMA wal_checkpoint(TRUNCATE)` (si aplica) + leer `TempDbPath` + `_ctx.ReEncryptAndPersist(bytes)` + `MarkChangesPersisted()`, con manejo de error sin excepción hacia la UI (`Result`). Implementar `ConfigRepository.UpdateSubunidad` con `WHERE id = 1`. En `FormUsuarios`: sustituir `db.Execute("UPDATE config ...")` por el repo y llamar `session.Persist()` tras `InsertUsuario`/`UpdateUsuario`/`DeleteUsuario`, mostrando error si la persistencia falla. Sin SQL en el code-behind.
- **validation**: `UsuariosPersistenciaTests` en GREEN y `Full` sin regresiones.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~UsuariosPersistenciaTests"`
- **evidence**: `_evidence/20261027-130512`
- **status**: Completed
- **log**: Implementado AcademicDbSession.Persist() con checkpoint WAL, lectura de bytes desencriptados y ReEncryptAndPersist retornando Result. Implementado ConfigRepository.UpdateSubunidad con soporte a IDbConnection y EncryptedSQLiteContext. Eliminado SQL directo en FormUsuarios.cs sustituyéndolo por ConfigRepository.UpdateSubunidad y llamadas a _session.Persist() tras guardado y eliminación. UsuariosPersistenciaTests en VERDE (1 pasado, 0 fallos).
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Context/AcademicDbSession.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/ConfigRepository.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`

### T4: Identidad determinista y lectura sin duplicados (implementer)
- **depends_on**: [T1]
- **location**: `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Sql/usuarios.sql`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Sql/usuarios.sql`
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IAcademicRepository.cs`
- **description**:
  - **(a) Lectura**: `GetUsuarios` con `ROW_NUMBER() OVER (PARTITION BY LOWER(TRIM(nombreusuario)) ORDER BY (firma no vacía) DESC, (perfil distinto de 'Academica') DESC, updated_at DESC, id)` filtrando `rn = 1` ⇒ una fila por usuario.
  - **(b) Escritura**: `InsertUsuario` con `id = "usr_" + lower(trim(nombreusuario))` (determinista) e inserción idempotente (`WHERE NOT EXISTS` por nombreusuario, manteniendo el mensaje de "ya registrado").
  - **(c) Seeds**: fijar el mismo `id` determinista en los 10 `INSERT INTO usuarios` de `usuarios.sql` (el `WHERE NOT EXISTS` por `nombreusuario` se mantiene).
  - Sin tocar `Consorcio`/permisos ni cambios de esquema.
- **validation**: `UsuariosDuplicadosTests` GREEN y `Full` sin regresiones.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~UsuariosDuplicadosTests"`
- **evidence**: `_evidence/20261027-130022`
- **status**: Completed
- **log**: Implementada lectura de usuarios con ROW_NUMBER deduplicando por nombreusuario con prioridad por firma/perfil no genérico/updated_at. Implementado InsertUsuario determinista con id='usr_' + lower(trim(nombreusuario)). Actualizados los 10 registros de semillas en usuarios.sql con sus respectivos IDs deterministas. UsuariosDuplicadosTests en VERDE (2 pasados, 0 fallos).
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Sql/usuarios.sql`

### T5: Resolver de subunidad determinista (integrator, L0 Core/Sync)
- **depends_on**: [T2]
- **location**: `ConversorPDF.XLL/Core/Sync/SubunidadResolver.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Core/Sync/SubunidadResolver.cs`
- **description**: añadir `ORDER BY activo DESC, updated_at DESC, id ASC` al `SELECT ... LIMIT 1` de `usuarios` y registrar en log los candidatos cuando haya más de uno (`WARNING` con `id|subunidad|updated_at`). Mantener el fallback a `config` y el comportamiento con tabla ausente.
- **validation**: `SubunidadResolverTests` GREEN (incluye caso con duplicados y con `updated_at` nulo).
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~SubunidadResolverTests"`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

### T6: Mover usuario entre vaults al cambiar de subunidad (implementer, condicional a T2)
- **depends_on**: [T2, T3, T5]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs`, `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`
  - `ConversorPDF.Tests/Integration/Usuarios/MoverUsuarioSubunidadTests.cs`
- **description**: si el test E2E de T2 demuestra que el cambio de subunidad del usuario propio debe viajar al vault destino: (1) upsert del usuario en `AcademicData_<DESTINO>.db.enc` vía `EncryptedSQLiteContext` (crear/aplicar esquema si no existe), (2) borrado en el vault origen para generar tombstone CRDT, (3) `Persist()` en ambos contextos, (4) si es el usuario actual, actualizar `config.subunidad` en ambos y mantener el aviso de reinicio. Si T2 **no** lo demuestra: no implementar, registrar `Ruling` y cerrar la tarea como `done` con la evidencia del descarte.
- **validation**: `MoverUsuarioSubunidadTests` GREEN (fila en destino, ausente en origen, ambos vaults persistidos) o `Ruling` documentado.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~MoverUsuarioSubunidad"`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

### T7: Migración V79 — normalizar IDs y deduplicar (integrator, L0)
- **depends_on**: [T4]
- **location**: `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV79_NormalizeUsuariosIds.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV79_NormalizeUsuariosIds.cs` (nuevo)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **description**: migración **idempotente** (versión 79; la máxima actual es V78) que: (1) hace copia de seguridad del vault (`*.bak`) antes de mutar; (2) deduplica por `LOWER(TRIM(nombreusuario))` conservando la fila más rica (firma > perfil no genérico > `updated_at` > `id`); (3) reescribe `id` a `usr_<lower(trim(nombreusuario))>` con `DELETE`+`INSERT` para que CR-SQLite genere tombstones propagables; (4) ejecuta `PRAGMA integrity_check` y registra el resultado. Log de filas normalizadas/eliminadas.
- **validation**: `verify.ps1 -Mode Build` exit 0 + `UsuariosDuplicadosTests` verde sobre una DB migrada + `integrity_check` = `ok`.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

### T8: Dedupe idempotente tras el merge de sync (integrator, L0 Core/Sync)
- **depends_on**: [T7]
- **location**: `ConversorPDF.XLL/Core/Sync/SyncInitializer.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Core/Sync/SyncInitializer.cs`
- **description**: extraer la deduplicación de `MigrationV41` a un método privado idempotente y ejecutarlo **después** de aplicar merges/pulls y **antes** de la autosanación de triggers, para que un pull de otro equipo no reintroduzca duplicados hasta el siguiente arranque. Log con `filas eliminadas`. No cambiar el orden del resto del pipeline de sync.
- **validation**: `UsuariosDuplicadosTests` (caso "pull reintroduce duplicado") verde + `Full` sin regresiones.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

### T9: Gate final, revisión y documentación (build-verifier + task-reviewer)
- **depends_on**: [T1..T8]
- **location**: `_evidence/` (solo lectura de código)
- **files_lock**: []
- **description**: ejecutar la suite completa en serial y revisar: sin SQL en code-behind, `Persist()` llamado en todas las rutas de escritura, sin Interop/WinForms en Application, sin C# 12, un tipo por archivo, `files edited/created` ⊆ `files_lock`, y lectura de usuarios sin duplicados. Escribir `walkthrough.md` + `docs/articles/changes/2026-10-27-usuarios-duplicados-subunidad.md`.
- **validation**: `verify.ps1 -Mode Full` exit 0, 0 regresiones, suite no reducida (939); veredicto `OK` del `task-reviewer`.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- **evidence**:
- **status**: pending
- **log**:
- **files edited/created**:

---

## 6. Olas de ejecución

| Wave | Tasks | Can start when |
|---|---|---|
| 1 | T1 (integrator, L0) | preflight `0` |
| 2 | T2, T4 | T1 completo (leases disjuntos: tests vs repositorio/seeds) |
| 3 | T3, T5, T7 | T3←T2, T5←T2, T7←T4 (tres leases disjuntos) |
| 4 | T6, T8 | T6←T2,T3,T5; T8←T7 |
| 5 | T9 | todo lo anterior |

Solapes evitados: `AcademicRepository.cs` solo en T4; `FormUsuarios.cs` en T1/T3/T6 (olas distintas); `csproj` en T1 y T7 (olas distintas); `SyncInitializer.cs` solo en T8; `SubunidadResolver.cs` solo en T5.

## 7. Estrategia de pruebas

- `Fast` por clase durante el ciclo; `Full` al cerrar cada tarea y en T9.
- La prueba E2E de persistencia usa **vault temporal** y un ensamblado cuyo nombre contenga `Tests` (el `Dispose` de `EncryptedSQLiteContext` detecta modo test y no avisa de pérdida de datos).
- Prohibido: debilitar asserts, subir timeouts, excluir tests con `--filter` fuera del SUT, o "arreglar" el test en vez del código.
- Evidencia en `_evidence/<timestamp>/summary.md`; no se commitea.

## 8. Riesgos y mitigaciones

| Riesgo | Mitigación |
|---|---|
| Reescritura de IDs en tabla CRDT rompe sync | Backup del vault + `PRAGMA integrity_check` + dedupe de lectura (la UI no muestra duplicados aunque la migración falle) + stop condition con `Ruling` |
| Doble vault abierto (locks del temp compartido) | Abrir destino en contexto separado, cerrar antes de persistir, reintentos acotados, nunca en paralelo con el gate |
| Cambio de constructor de `FormUsuarios` | Solo hay un call-site (`RibbonController.cs:1256`), verificado por grep |
| `config` sin `WHERE` afecta varias filas | `UPDATE config SET subunidad=@S WHERE id = 1` en T3 |
| El pull de sync reintroduce duplicados | T8 (dedupe post-merge) + identidad determinista (T4/T7) |
| Tests de UI/STA | Los tests nuevos son de persistencia/servicio (sin WinForms), no requieren STA |

## 9. Criterios de aceptación (manual + automático)

1. **Duplicados**: con dos filas del mismo `nombreusuario`, el grid muestra **una**; tras reiniciar Excel sigue mostrando una.
2. **Persistencia**: cambiar subunidad y guardar ⇒ cerrar Excel ⇒ reabrir ⇒ la fila conserva la subunidad nueva.
3. **Config coherente**: `config.subunidad` (id=1) coincide con la del usuario actual y sobrevive al reinicio.
4. **Sync**: un pull que reintroduce un duplicado no queda visible ni persistente tras el arranque.
5. **Regresión**: `verify.ps1 -Mode Full` → `PASS`, 939+ pruebas, 0 fallos, suite no reducida.

## 10. Fuera de alcance

- Rediseñar la partición por subunidad (se mantiene un vault por subunidad).
- Limpiar tombstones de CR-SQLite o cambiar la política de sync.
- Reasignar permisos/perfiles o tocar `permisos`.
- Cambiar el formato del seed `usuarios.sql` (se mantiene aditivo con `WHERE NOT EXISTS`).

---

## 11. Autoevaluación del plan (checklist de 12 puntos)

| # | Punto | Resultado |
|---|---|---|
| 1 | Dependencias completas | ✅ T1..T9 encadenadas; T9 depende de todas |
| 2 | Orden ejecutable | ✅ 5 olas; cada tarea compila al terminar (T1 crea esqueletos) |
| 3 | Ninguna tarea mezcla L0 sin `integrator` | ✅ L0 solo en T1 (wiring/esqueletos), T5 (`SubunidadResolver`), T7 (migración+csproj), T8 (`SyncInitializer`) |
| 4 | `files_lock` disjuntos por ola | ✅ verificado ola por ola (§6) |
| 5 | `verification` ejecutable | ✅ todos usan `scripts/verify.ps1` con filtros reales |
| 6 | `validation` falsable | ✅ aserciones concretas por test + exit codes |
| 7 | El RED compila | ✅ T1 crea `Persist()`, `UpdateSubunidad` y servicio stub antes de T2 |
| 8 | Migración/versión asignada por `integrator` | ✅ V79 (máxima actual V78), solo en T7 |
| 9 | Ninguna tarea cruza slices | ✅ todo en `Academic` + `Core/Sync` (declarados L0 en el envelope) |
| 10 | Presupuesto de contexto | ✅ sin pack aplicable (no existe carpeta `Usuarios`); inventario de archivos incluido en §2/§5 |
| 11 | Criterio de aceptación global | ✅ §9 (5 criterios, 2 manuales + 3 automáticos) |
| 12 | Riesgos con mitigación | ✅ §8, incluida la operación CRDT destructiva |

**Correcciones aplicadas tras la autoevaluación:** (a) se fusionó el antiguo "wiring en RibbonController" dentro de T1 para no tocar L0 dos veces; (b) se separó `GetUsuarios`/`InsertUsuario`/seeds en una sola tarea (T4) porque comparten archivo; (c) se hizo T6 condicional con `Ruling` para no construir un mecanismo que T2 podría no justificar; (d) se añadió backup + `integrity_check` a T7 por ser operación destructiva sobre CRDT.
