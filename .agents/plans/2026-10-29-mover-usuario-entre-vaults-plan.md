# Plan: Mover usuario entre vaults al cambiar de subunidad (MUV)

**Generado**: 2026-10-29 · **Autor**: Cline/Swarm-Planner · **Estado**: approved

## Execution envelope
- branch: refactor/academic-modularization
- workspace: árbol real bajo OneDrive (456 `.cs` XLL / 143 `.cs` Tests), validado con `-AllowOneDrive`; para la Fase B se recomienda copia fuera de OneDrive (`C:\dev\ConversorPDF`)
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <workspace> -Plan .agents/plans/2026-10-29-mover-usuario-entre-vaults-plan.md -AllowOneDrive
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: 952 · medido 2026-10-27 en C:\dev\ConversorPDF; el preflight de este plan ya reporta TOTAL=952)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado (un único build-verifier a la vez; Excel-DNA bloquea bin/)
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0:
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Core/Helpers/App.cs
  - ConversorPDF.XLL/**/Migrations/MigrationV*.cs
  - ConversorPDF.XLL/**/Persistence/SchemaMigrator.cs
  - ConversorPDF.XLL/Ribbon/RibbonController.cs
  - ConversorPDF.XLL/Core/Sync/SyncInitializer.cs
  - Composition Root / IoC Container
  - docs/toc.yml
- stop_conditions:
  - `DELETE` de filas en `usuarios` dentro del flujo de movimiento (rompe CRDT: el tombstone puede borrar la fila del vault destino)
  - `UPDATE config SET subunidad` cuando el usuario editado no es el usuario de Windows actual
  - Acción destructiva no reversible sin respaldo
  - Intento de push, merge o publicación a remoto
  - Plan irrecuperable tras 3 rondas de fix
- ledger_owner: swarm-parallel-task
- no_push: true

---

## 1. Objetivo

Que **cambiar la subunidad de un usuario en Gestión de Usuarios lo mueva realmente de vault** (fila correcta en `AcademicData_<DESTINO>.db.enc`) y que `config.subunidad` (id=1) quede coherente **en el vault del dispositivo**, sin romper la replicación CRDT ni el arranque ya validado por `R-20261027-T6-02`.

Requisitos del usuario, textuales:

1. «Mover usuario entre vaults al cambiar de subunidad».
2. «Al cambiar subunidad en FormUsuarios debe cambiar o actualizar en FormConfig la subunidad».

## 2. Diagnóstico (evidencia en código, no suposiciones)

| # | Defecto | Evidencia |
|---|---|---|
| **D1** | El servicio de movimiento es un **no-op** | `MoverUsuarioSubunidadService.cs:13-21` — solo valida vacíos y devuelve `Result.Success()` |
| **D2** | **Sin call-sites**: nunca se invoca | grep `IMoverUsuarioSubunidadService` en `.cs` → 0 usos; solo 3 hits en `.csproj:353-354` (código muerto) |
| **D3** | `FormUsuarios` escribe **solo en el vault activo**; el destino nunca recibe la fila | `FormUsuarios.cs:449` (`InsertUsuario`), `:456` (`UpdateUsuario`), `:495` (`_session.Persist()`) |
| **D4** | `config.subunidad` se actualiza **solo** si el editado es `Environment.UserName` | `FormUsuarios.cs:464-473` |
| **D5** | El vault destino solo recibe al **usuario de Windows**, en el arranque | `SyncContext.cs:430-466` (`vaultChanged` → `UsuarioVaultProvisioner.EnsureUsuarioEnVault`) |
| **D6** | `UsuarioVaultProvisioner` fuerza `perfil='Academica'` y `activo=1` | `UsuarioVaultProvisioner.cs:48-64` — no preserva perfil/firma/activo reales |
| **D7** | `usuarios` **y** `config` son **tablas CRDT replicadas** ⇒ un `DELETE` genera tombstone propagable | `SyncInitializer.cs:21-26` (`CrdtTables` incluye `"usuarios"`, `"config"`); `SubunidadSyncService.cs:73` |
| **D8** | `FormConfig` deja el combo **read-only** y no se recarga si ya está abierto | `FormConfig.cs:100` (`cmbSubunidad.Enabled = false`); carga solo en el ctor (`:92-150`) |

## 2.1 Auditoría de los artefactos previos del plan 2026-10-27

### T6 — «Mover usuario entre vaults al cambiar de subunidad (implementer, condicional a T2)»

Verificado verbatim en `.agents/plans/2026-10-27-usuarios-duplicados-subunidad-plan.md:185-199`:

| Campo | Valor real | Hallazgo |
|---|---|---|
| `status` | `completed` | **Engañoso**: figura cerrada sin código funcional |
| `description` | «…Si T2 **no** lo demuestra: no implementar, registrar `Ruling` y cerrar la tarea como `done`» | La cláusula de escape se usó |
| `evidence` | `Ruling R-20261027-T6-01` | Apunta a un Ruling, no a evidencia de ejecución |
| `ruling_vigente` | `R-20261027-T6-02` (ver §12) | `T6-01` **supersedido** por la auditoría independiente |
| `log` | «persistir el vault actual … es 100% suficiente» | **Refutado** por §12: solo probaba reabrir el **mismo** vault |

**Conclusión de T6:** el descarte fue legítimo como decisión de diseño, pero el `status: completed` oculta una **deuda técnica explícita**: hoy `Mover()` no mueve nada y no tiene consumidores (D1/D2). Este plan **salda esa deuda** en lugar de reabrir el plan anterior (regla de oro de `.agents/plans/README.md` §5.4).

### T12 — «`UpdateUsuario` sella `updated_at`; `GetUsuarios`/`DeduplicateUsuarios` prefieren `usr_`»

Re-verificado en código: **3/3 correcto, sin hallazgo**.

| Afirmación de T12 | Evidencia |
|---|---|
| `UpdateUsuario` sella `updated_at = datetime('now')` | `AcademicRepository.cs:1478` ✅ |
| `GetUsuarios` prefiere identidad determinista | `AcademicRepository.cs:1397` — `CASE WHEN substr(id,1,4)='usr_' THEN 0 ELSE 1 END ASC` dentro de `ROW_NUMBER()` ✅ |
| `DeduplicateUsuarios` prefiere identidad determinista | `SyncInitializer.cs:912` — misma cláusula ✅ |
| Identidad determinista en alta | `AcademicRepository.cs:1456` — `usuario.Id = "usr_" + normalizedUser` ✅ |

**Matiz de ledger:** T12 no es tarea del plan base (T1–T9) sino del T-fix (§12, T10–T13); por eso no tiene bloque propio, solo la fila en la tabla del §12 con la evidencia «tests H2/H3 en verde». Consistente.

### §12 — T-fix de la auditoría independiente

| Verificación | Resultado |
|---|---|
| ¿Existe `Ruling` vigente? | ✅ `R-20261027-T6-02`: **sí** debe provisionarse la fila del usuario en el vault destino al conmutar |
| ¿El provisioner está implementado? | ✅ `UsuarioVaultProvisioner.EnsureUsuarioEnVault` (Core/Sync, `internal static`) |
| ¿Se invoca en ambos caminos? | ✅ `SyncContext.cs:438` `if (newlyRegistered \|\| vaultChanged)` → `:457` |
| ¿Cobertura honesta? | ✅ §12 la declara: provisioner con tests (4 casos en `UsuarioVaultProvisionerTests`); invocación desde `SyncContext` **solo por revisión**, sin E2E (método modal con `FormRegistroSubunidad`) |
| Hueco que deja abierto | 🔴 El provisioner cubre solo al **usuario de Windows**; mover a un **tercero** entre vaults sigue sin existir (D5/D6) |

## 3. Decisiones de diseño (Rulings)

### R-20261029-MUV-01 — El movimiento se hace con `UPDATE` + `UPSERT`, nunca con `DELETE`

`usuarios` es tabla CRDT (`SyncInitializer.cs:21-26`). Un `DELETE` en el vault origen genera un **tombstone** que se propaga y puede eliminar la fila recién insertada en el destino (misma PK determinista `usr_<nombre>`). Por tanto:

1. **Origen**: se mantiene la escritura existente `AcademicRepository.UpdateUsuario` (subunidad nueva + `updated_at`), **sin borrar**.
2. **Destino**: upsert idempotente de la fila completa (nombre, subunidad, perfil, activo, firma, `updated_at`).
3. Ambos vaults se persisten; no hay tombstones en el flujo.

Consecuencia funcional: el usuario aparece en el vault destino de inmediato y, si arranca en el origen, `SubunidadResolver` devuelve la subunidad nueva y conmuta al vault destino (mecanismo ya validado por `R-20261027-T6-02`). Coste si la decisión es errónea: una fila duplicada por vault, corregible por `DeduplicateUsuarios` (idempotente, `SyncInitializer.cs:912`).

### R-20261029-MUV-02 — `config.subunidad` es del **dispositivo**, no de la persona

`config` también es CRDT y replica `subunidad`: si al editar a un tercero se cambiara `config.subunidad`, **todos los dispositivos** de la réplica saltarían de vault (regresión del bug «la UI abre con territorial incorrecta»). Por tanto `config.subunidad` (id=1) se actualiza **solo** cuando el usuario editado es el usuario de Windows actual (comparando nombre nuevo **y** anterior, como ya hace `FormUsuarios.cs:464-466`). Criterio falsable: test unitario `Mover_WhenUsuarioNoEsElActual_ShouldNotTouchConfig`.

### R-20261029-MUV-03 — No se reutiliza `UsuarioVaultProvisioner` para el alta en destino

`EnsureUsuarioEnVault` fija `perfil='Academica'` y `activo=1` (D6), lo que **degradaría** el usuario movido. El upsert del movimiento se implementa en `UsuarioVaultMover` (Infraestructura) preservando perfil, activo y firma. `UsuarioVaultProvisioner` queda intacto (no se toca Core/Sync → menos riesgo de regresión en el arranque).

### R-20261029-MUV-04 — Seam de test sin nuevos tipos

`UsuarioVaultMover` expone un ctor público sin parámetros (usa `SqliteHelper`) y un ctor `internal` que recibe tres delegados (`resolvePath`, `openVault`, `persist`). Así el E2E inyecta vaults temporales **sin** añadir interfaces nuevas ni carpetas nuevas al slice.

---

## 4. Dependency graph

```text
T1 (integrator, L0: contratos + esqueletos + csproj)
 ├── T2 (RED: 1 unit + 2 integration)
 │    ├── T3 (implementer: UsuarioVaultMover)
 │    └── T4 (implementer: MoverUsuarioSubunidadService)
 │         └── T5 (implementer: FormUsuarios + FormConfig)
 │              └── T6 (gate: Full + task-reviewer + docs)
```

## 5. Tasks

### T1: Contratos, DTOs, esqueletos y registro en csproj (integrator, L0)
- **id**: T1
- **depends_on**: []
- **location**: `ConversorPDF.XLL/Academic/Application/{Interfaces,DTOs,Services}`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Interfaces/IUsuarioVaultMover.cs` (nuevo)
  - `ConversorPDF.XLL/Academic/Application/Interfaces/IMoverUsuarioSubunidadService.cs` (firma ampliada)
  - `ConversorPDF.XLL/Academic/Application/DTOs/MovimientoVaultResultado.cs` (nuevo)
  - `ConversorPDF.XLL/Academic/Application/DTOs/MovimientoUsuarioResultado.cs` (nuevo)
  - `ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs` (ctor + esqueleto)
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/UsuarioVaultMover.cs` (nuevo, esqueleto)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **description**:
  - Puerto `IUsuarioVaultMover`: `Result<MovimientoVaultResultado> Mover(Usuario usuario, string subunidadDestino)`.
  - Servicio: `Result<MovimientoUsuarioResultado> Mover(Usuario usuario, string subunidadDestino, string usuarioActual)` con ctor `(IUsuarioVaultMover vaultMover, IConfigRepository configRepository)`. El formulario ya tiene el `Usuario` completo en memoria, así que el servicio no lee el vault (menos I/O y tests más baratos).
  - DTOs: `MovimientoVaultResultado { string SubunidadDestino; string RutaVaultDestino; bool Insertado; bool Actualizado; }`; `MovimientoUsuarioResultado { string UsuarioId; string SubunidadDestino; bool ConfigActualizada; bool RequiereReinicio; }`.
  - Esqueleto de `UsuarioVaultMover` con los dos ctors de `R-20261029-MUV-04` y `Mover` devolviendo `Result.Failure("Mover entre vaults no implementado")`; el servicio delega y devuelve `Failure` (RED compilable, sin lógica).
  - Registrar los 4 archivos nuevos en `<Compile Include>` **sin duplicar** (evitar CS2002) y un tipo por archivo.
- **validation**: `verify.ps1 -Mode Build` exit 0 y `grep` de los 4 archivos nuevos en el `.csproj` (exactamente 1 vez cada uno).
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build`
- **evidence**: `_evidence/20261029-104821`
- **status**: completed
- **log**: Contratos IUsuarioVaultMover, ampliación de IMoverUsuarioSubunidadService, DTOs MovimientoVaultResultado y MovimientoUsuarioResultado, esqueletos de MoverUsuarioSubunidadService y UsuarioVaultMover creados y registrados en ConversorPDF.XLL.csproj. Compilación en verde sin errores (verify.ps1 -Mode Build exit 0).
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/Application/Interfaces/IUsuarioVaultMover.cs`
  - `ConversorPDF.XLL/Academic/Application/Interfaces/IMoverUsuarioSubunidadService.cs`
  - `ConversorPDF.XLL/Academic/Application/DTOs/MovimientoVaultResultado.cs`
  - `ConversorPDF.XLL/Academic/Application/DTOs/MovimientoUsuarioResultado.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/UsuarioVaultMover.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`

### T2: RED — movimiento, orquestación y FormConfig (tdd-red-writer)
- **id**: T2
- **depends_on**: [T1]
- **location**: `ConversorPDF.Tests/Unit/Usuarios/`, `ConversorPDF.Tests/Integration/Usuarios/`
- **files_lock**:
  - `ConversorPDF.Tests/Unit/Usuarios/MoverUsuarioSubunidadServiceTests.cs` (nuevo)
  - `ConversorPDF.Tests/Integration/Usuarios/UsuarioVaultMoverTests.cs` (nuevo)
  - `ConversorPDF.Tests/Integration/Usuarios/MoverUsuarioEntreVaultsE2ETests.cs` (nuevo)
- **description**:
  - **(a) Unit, fakes en memoria** (`FakeUsuarioVaultMover`, `FakeConfigRepository` con `Llamadas`):
    - `Mover_WhenUsuarioNoEsElActual_ShouldNotTouchConfig` (R-20261029-MUV-02)
    - `Mover_WhenUsuarioEsElActual_ShouldUpdateConfigAndRequireRestart`
    - `Mover_WhenVaultMoverFalla_ShouldPropagateFailureAndNotTouchConfig`
    - `Mover_WhenSubunidadDestinoVacia_ShouldFail`
    - `Mover_WhenUsuarioIdVacio_ShouldFail`
  - **(b) Integration SQL** (`:memory:`, esquema `usuarios` copiado de `UsuarioVaultProvisionerTests`):
    - `Mover_WhenUsuarioExisteEnOrigen_ShouldPreservePerfilActivoYFirma` (R-MUV-03)
    - `Mover_WhenSeInvocaDosVeces_ShouldSerIdempotente`
    - `Mover_WhenUsuarioNoExisteEnDestino_ShouldInsertarConIdDeterminista`
  - **(c) E2E dos vaults reales** (patrón de `UsuariosPersistenciaTests`: temp dir + `EncryptedSQLiteContext` + `SecureString`), inyectando los delegados del ctor `internal`:
    - `Mover_WhenCambiaSubunidad_ShouldPersistirFilaEnVaultDestinoYReabrirla`
    - `Mover_WhenRutaDestinoIgualAOrigen_ShouldSerNoOpSinPerdida`
  - xUnit + FluentAssertions 6.12; nombres `{Accion}_When_{Condicion}_Should_{Resultado}`; **sin** borrar filas para probar el movimiento (R-MUV-01).
- **validation**: compilan y fallan **por la razón prevista** (`Failure("no implementado")` o aserción concreta), nunca por error de compilación.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~MoverUsuarioSubunidadServiceTests|FullyQualifiedName~UsuarioVaultMoverTests|FullyQualifiedName~MoverUsuarioEntreVaultsE2ETests"`
- **evidence**: `_evidence/20261029-105218`
- **status**: completed
- **log**: Fase RED certificada. 12 pruebas ejecutadas (Unit, Integration SQL y E2E), 4 casos pasaron (guards vacíos), 8 fallaron por la razón esperada (esqueletos devuelven Failure de no implementado). Compilación correcta.
- **files edited/created**:
  - `ConversorPDF.Tests/Unit/Usuarios/MoverUsuarioSubunidadServiceTests.cs`
  - `ConversorPDF.Tests/Integration/Usuarios/UsuarioVaultMoverTests.cs`
  - `ConversorPDF.Tests/Integration/Usuarios/MoverUsuarioEntreVaultsE2ETests.cs`



### T3: Implementar el movimiento entre vaults (implementer)
- **id**: T3
- **depends_on**: [T2]
- **location**: `ConversorPDF.XLL/Academic/Infrastructure/Persistence/UsuarioVaultMover.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/UsuarioVaultMover.cs`
- **description**:
  - `Mover(usuario, subunidadDestino)`: normalizar destino; resolver `destino = resolvePath(subunidad)` y `origen = SqliteHelper.GetVaultPath()`; si son iguales (OrdinalIgnoreCase) ⇒ `Result.Success` con `Insertado=false, Actualizado=false` (no-op, R-MUV-04).
  - Abrir el vault destino con `openVault(destino)` en `using`, `Open()`, y ejecutar `UpsertEnVault` (método `internal` que recibe `IDbConnection`, para el test SQL de `:memory:`).
  - `UpsertEnVault`: `INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, perfil, activo, firma, updated_at) VALUES (...) ON CONFLICT(id) DO UPDATE SET nombre/subunidad/perfil/activo/firma/updated_at` con `id = "usr_" + lower(trim(nombreusuario))`; preserva perfil, activo y firma (R-MUV-03); sella `updated_at = datetime('now')`; devuelve si fue INSERT o UPDATE vía `SELECT COUNT(1)` previo.
  - Persistir el destino: `CloseConnectionResiliently()` → `EncryptedSQLiteContext.LeerBytesConRetrySync(ctx.TempDbPath)` → `persist(bytes, destino)` → `MarkChangesPersisted()`.
  - **Prohibido**: `DELETE FROM usuarios` (R-MUV-01) y `SetTestVaultPath` (el vault activo solo cambia en el arranque).
  - Todo error se envuelve en `Result.Failure` con contexto (sin excepciones hacia la UI).
- **validation**: `UsuarioVaultMoverTests` y `MoverUsuarioEntreVaultsE2ETests` en verde; `Full` sin regresiones.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~UsuarioVaultMoverTests|FullyQualifiedName~MoverUsuarioEntreVaultsE2ETests"`
- **evidence**: `_evidence/20261029-105529`
- **status**: completed
- **log**: Implementado UsuarioVaultMover con UpsertEnVault idempotente preservando perfil, activo y firma. Persistencia atómica de bytes y re-apertura en destino validadas con 5 tests en verde (3 integration + 2 E2E).
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/UsuarioVaultMover.cs`

### T4: Implementar la orquestación y la política de config (implementer)
- **id**: T4
- **depends_on**: [T2]
- **location**: `ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs`
- **description**:
  - Guards: `usuario == null`, `NombreUsuario` vacío, `subunidadDestino` vacía ⇒ `Result.Failure` con mensaje de negocio.
  - Delegar en `_vaultMover.Mover(usuario, destinoNormalizado)`; si falla ⇒ propagar **sin** tocar `config`.
  - `esUsuarioActual = string.Equals(usuario.NombreUsuario?.Trim(), usuarioActual?.Trim(), StringComparison.OrdinalIgnoreCase)`.
  - Si `esUsuarioActual`: `_configRepository.UpdateSubunidad(destino)`; si falla ⇒ `Result.Failure`; si no ⇒ `ConfigActualizada = true`, `RequiereReinicio = true` (**R-MUV-02**). Si no es el actual ⇒ `ConfigActualizada = false`, `RequiereReinicio = false`.
  - Sin `System.Data`, sin WinForms, sin Interop: solo `Result` y los dos puertos (regla de capas).
- **validation**: `MoverUsuarioSubunidadServiceTests` en verde (5 casos), incluido `Mover_WhenUsuarioNoEsElActual_ShouldNotTouchConfig`.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~MoverUsuarioSubunidadServiceTests"`
- **evidence**: `_evidence/20261029-105657`
- **status**: completed
- **log**: Implementado MoverUsuarioSubunidadService con política R-MUV-02 (config.subunidad solo se actualiza si el editado es el usuario actual de Windows), validación de negocio y delegación en IUsuarioVaultMover. 7 unit tests en verde.
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs`

### T5: Cablear el formulario y recargar FormConfig (implementer)
- **id**: T5
- **depends_on**: [T3, T4]
- **location**: `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`, `ConversorPDF.XLL/Academic/UI/Forms/FormConfig.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormConfig.cs`
- **description**:
  - En `btnGuardar`: tras `UpdateUsuario`/`InsertUsuario`, si `subunidadCambiada`, construir `new MoverUsuarioSubunidadService(new UsuarioVaultMover(), new ConfigRepository(db))` y llamar `Mover(u, u.Subunidad, Environment.UserName)`; si `IsFailure` ⇒ `MessageBox` + `return` (no se persiste nada más).
  - Sustituir el bloque `FormUsuarios.cs:461-492` por la lectura de `MovimientoUsuarioResultado.RequiereReinicio` (el servicio encapsula la política de `config`); se conserva el mensaje de «Reinicio Requerido» y `DialogResult.Abort`.
  - Orden de persistencia: `mover` (persiste el destino) → servicio actualiza `config` en la conexión activa → `_session.Persist()` (persiste origen: usuario + config).
  - `FormConfig`: extraer la carga del ctor a `private void CargarConfiguracion()` y añadir `public void RecargarConfiguracion()`; en `FormUsuarios`, tras mover al usuario actual, invocar `RecargarConfiguracion()` sobre la instancia abierta (`Application.OpenForms`) para que el combo no quede obsoleto (D8).
  - Sin SQL nuevo en el code-behind: solo composición manual (igual que `new ConfigRepository(db)` ya existente) y mensajes.
- **validation**: `Full` exit 0 sin regresiones; `grep` de `UPDATE config`/`Execute(` en `FormUsuarios.cs` sin nuevas apariciones; `RecargarConfiguracion` existe y es invocado.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~Usuarios"`
- **evidence**: `_evidence/20261029-110002`
- **status**: completed
- **log**: Cableado MoverUsuarioSubunidadService en FormUsuarios.btnGuardar con persistencia secuencial (destino primero, luego sesión de origen), reinicio requerido condicional a usuario actual, y recarga de FormConfig abierto vía RecargarConfiguracion(). 25 tests de Usuarios superados.
- **files edited/created**:
  - `ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormConfig.cs`


### T6: Gate final, revisión de arquitectura y documentación (build-verifier + task-reviewer)
- **id**: T6
- **depends_on**: [T5]
- **location**: `_evidence/`, `docs/articles/changes/`, `walkthrough.md`
- **files_lock**:
  - `docs/articles/changes/2026-10-29-mover-usuario-entre-vaults.md` (nuevo)
  - `walkthrough.md`
- **description**:
  - Ejecutar `-Mode Full` en serial y certificar `TOTAL >= 952` (baseline), 0 regresiones, `SUITE_SHRUNK` ausente.
  - Revisión `task-reviewer` en 2 etapas: (1) spec compliance de AC1–AC5; (2) calidad: sin `System.Data` en `Application`, sin Interop/WinForms en `Application`, sin C# 12, un tipo por archivo, `files edited/created` ⊆ `files_lock`, **sin `DELETE FROM usuarios`** en el flujo, `config.subunidad` solo para el usuario actual.
  - Documentar en `docs/articles/changes/2026-10-29-mover-usuario-entre-vaults.md` (Purpose · Changes · How to use · Test coverage) y `walkthrough.md`.
- **validation**: `verify.ps1 -Mode Full` exit 0 con `TOTAL >= 952`; veredicto `OK` del `task-reviewer`; los 5 AC de §9 documentados como verificados o manuales.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- **evidence**: `_evidence/20261029-110127`
- **status**: completed
- **log**: Gate final superado con 964 tests aprobados (TOTAL >= 952 baseline, 0 errores, 0 regresiones). Revisión task-reviewer de 2 etapas OK (AC1-AC5 conformes, sin System.Data en Application, sin DELETE en flujo, un tipo por archivo, C# 12 ausente). Documentación técnica generada en docs/articles/changes/2026-10-29-mover-usuario-entre-vaults.md y actualizada en walkthrough.md.
- **files edited/created**:
  - `docs/articles/changes/2026-10-29-mover-usuario-entre-vaults.md`
  - `walkthrough.md`

---

## 6. Parallel execution groups

| Wave | Tasks | Can start when |
|---|---|---|
| 1 | T1 (integrator, L0) | preflight `0` (`-AllowOneDrive` en este árbol) |
| 2 | T2 | T1 completado |
| 3 | T3, T4 | T2 completado (leases disjuntos: infra vs application; 2 escritores = tope) |
| 4 | T5 | T3 y T4 completados |
| 5 | T6 | T5 completado |

Solapes evitados: `ConversorPDF.XLL.csproj` solo en T1; `UsuarioVaultMover.cs` solo en T3; `MoverUsuarioSubunidadService.cs` solo en T1 y T4 (olas distintas); `FormUsuarios.cs` y `FormConfig.cs` solo en T5; tests solo en T2.

## 7. Testing strategy

- `Fast` por clase durante el ciclo RED→GREEN; `Full` en T5 y T6.
- **Unit** (sin I/O): política de `config` y propagación de fallos, con fakes de los dos puertos.
- **Integration SQL** (`:memory:`): el upsert preserva perfil/activo/firma y es idempotente.
- **E2E dos vaults reales** (temp dir + `EncryptedSQLiteContext` + `SecureString`, patrón `UsuariosPersistenciaTests`): mover y **reabrir** el destino; el `Dispose` de `EncryptedSQLiteContext` detecta el modo test por el nombre del ensamblado (`*Tests`).
- **Prohibido**: borrar filas para simular el movimiento, debilitar asserts, subir timeouts, excluir tests con `--filter` fuera del SUT.
- Evidencia en `_evidence/<timestamp>/summary.md`; no se commitea.

## 8. Risks & mitigations

| Riesgo | Mitigación |
|---|---|
| Tombstone CRDT borra la fila del destino | R-MUV-01: nunca `DELETE`; origen por `UPDATE`, destino por `UPSERT`; `stop_conditions` lo declara |
| `config.subunidad` replicada salta de vault en todos los dispositivos | R-MUV-02: solo se actualiza si el editado es el usuario de Windows; test dedicado |
| Doble vault abierto / locks WAL | destino en `using` con `CloseConnectionResiliently()` + `LeerBytesConRetrySync`; gate serializado (un solo `build-verifier`) |
| Vault destino inexistente | `SqliteHelper.GetContextForPath` lo crea y migra (`InitializeDatabase` + `EnsureSchemaUpToDate`); el E2E cubre el caso |
| Usuario movido queda sin perfil/firma | upsert completo en `UsuarioVaultMover` (no se reutiliza `UsuarioVaultProvisioner`, R-MUV-03) |
| `FormConfig` abierto muestra datos viejos | `RecargarConfiguracion()` invocado tras mover al usuario actual (D8) |
| Regresión en el arranque | `SyncContext.cs`, `SubunidadResolver.cs` y `UsuarioVaultProvisioner.cs` **no** se tocan (fuera de `files_lock`) |

## 9. Criterios de aceptación

| # | Criterio | Tipo |
|---|---|---|
| **AC1** | Cambiar la subunidad de **otro** usuario ⇒ su fila existe en `AcademicData_<DESTINO>.db.enc` con la subunidad nueva, conservando `perfil`, `activo` y `firma` | Automático (T2/T3) |
| **AC2** | La fila del **origen** queda con la subunidad nueva (sin desaparecer) y el vault origen se persiste | Automático (T2/T3) |
| **AC3** | Alcance del arranque: reabrir el add-in con la subunidad cambiada ⇒ el resolver devuelve la subunidad nueva y conmuta al vault destino, donde la fila ya existe | Manual (requiere Excel) |
| **AC4** | `config.subunidad` (id=1) cambia **solo** cuando el usuario editado es el usuario de Windows; `FormConfig` abierto se recarga | Automático (unit) + Manual (UI) |
| **AC5** | `verify.ps1 -Mode Full` ⇒ exit 0, `TOTAL >= 952`, 0 fallos, suite no reducida | Automático (T6) |

## 10. Fuera de alcance

- Cambiar la política de un vault por subunidad o el mecanismo de conmutación del arranque (`SyncContext`).
- Reescribir `UsuarioVaultProvisioner` o `SubunidadResolver`.
- Sincronización forzada inmediata tras el movimiento (el scheduler sigue su ciclo).
- Migraciones de esquema: este plan **no** crea `MigrationV*` (no hay cambio de DDL).
- Reasignar permisos/perfiles del usuario movido.

---

## 11. Autoevaluación del plan (checklist de 12 puntos)

| # | Punto | Resultado |
|---|---|---|
| 1 | Dependencias completas | ✅ T1←∅ · T2←T1 · T3←T2 · T4←T2 · T5←T3,T4 · T6←T5 |
| 2 | Orden ejecutable | ✅ 5 olas; T1 crea los esqueletos, así que el código compila en cada ola |
| 3 | Ninguna tarea mezcla L0 sin `integrator` | ✅ L0 (csproj) solo en T1 (`integrator`); `SyncContext`, `SubunidadResolver` y `UsuarioVaultProvisioner` **no** se tocan |
| 4 | `files_lock` disjuntos por ola | ✅ verificado en §6 (ola 3: infra vs application; olas 1-2-4-5 con un solo escritor) |
| 5 | Cada `verification` es un comando existente | ✅ `scripts/verify.ps1` con `-Mode Build/Fast/Full` y filtros reales (verificado: `verify.ps1:169` pasa `--filter` a `dotnet test`) |
| 6 | Cada `validation` es falsable | ✅ aserciones por test + exit codes + greps concretos |
| 7 | El RED compila | ✅ T1 crea `IUsuarioVaultMover`, DTOs y esqueletos que devuelven `Failure` antes de T2 |
| 8 | Migración/versión por `integrator` | ✅ N/A: **no hay migración** en este plan (sin cambio de DDL); si apareciera, sería tarea `integrator` nueva |
| 9 | Ninguna tarea cruza más de un slice | ✅ todo en el slice `Academic` (+ `Core/Sync` solo lectura) |
| 10 | Presupuesto de contexto ≤ 400 líneas | ⚠️ El context pack `academic-genetica.md` (138 líneas) **no cubre** este dominio. Se sustituye por un inventario acotado en §2 (8 defectos con archivo:línea) y §5 (`files_lock` por tarea). Recomendado generar `academic-sync-usuarios.md` antes de replicar este flujo |
| 11 | Criterio de aceptación global | ✅ §9 con 5 AC (3 automáticos, 2 manuales) |
| 12 | Riesgos con mitigación | ✅ §8 con 7 riesgos, incluido el tombstone CRDT y la replicación de `config` |

**Correcciones aplicadas tras la autoevaluación:**

- (a) La firma del servicio pasó de `Mover(usuarioId, …)` a `Mover(Usuario, …)`: el formulario ya tiene la entidad y así se evita una lectura extra al vault y un puerto más.
- (b) El E2E necesitaba rutas de vault inyectables: se resolvió con un **ctor `internal` de delegados** (`R-MUV-04`) en lugar de crear `IVaultAccess` (menos tipos, menos superficie).
- (c) Se descartó reutilizar `UsuarioVaultProvisioner` (`R-MUV-03`) porque fuerza `perfil='Academica'` y `activo=1`, lo que degradaría al usuario movido.
- (d) Se añadió el `stop_condition` del `DELETE` tras descubrir que `usuarios` **y** `config` son CRDT (`SyncInitializer.cs:21-26`).
- (e) Se incluyó `FormConfig.RecargarConfiguracion()` para cumplir el requisito 2 del usuario sin reabrir el formulario.

---

## 12. Handoff a Fase B

- **Handoff**: `.agents/plans/2026-10-29-mover-usuario-entre-vaults-handoff.md`.
- **Pre-flight exigido**: `preflight.ps1 … -AllowOneDrive` exit `0` (hoy devuelve `0` salvo el check `PLAN`, que este archivo resuelve) y `git status --porcelain` vacío.
- **Ejecución**: `/swarm-parallel-task .agents/plans/2026-10-29-mover-usuario-entre-vaults-plan.md` (o `executing-plans-local` si no hay `invoke_subagent`).
- **Cierre**: skill `swarm-verify` (Fase C) con `base` = `7381836` (HEAD actual).
- **Aprobación humana requerida** antes de la ola 1.

**Cobertura declarada desde ya (para no repetir el error de T6):** AC1, AC2, AC4(unit) y AC5 son automatizables y exigibles por gate; **AC3 y AC4(UI) requieren Excel abierto** y se validan por revisión + prueba manual registrada en el ledger. Si en la ola 3 el E2E de dos vaults resulta inviable en el host (locks/bloqueo de `EncryptedSQLiteContext`), la tarea **no** se cierra como `completed`: se registra `Ruling` y se degrada a cobertura por revisión, dejándolo escrito.


