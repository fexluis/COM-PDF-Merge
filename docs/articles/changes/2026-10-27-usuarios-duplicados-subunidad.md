# Normalización de Identidades de Usuarios, Persistencia del Vault y Resolución Determinista de Subunidad

## Purpose
Este cambio resuelve dos defectos críticos reportados en la gestión de usuarios y cambio de subunidad territorial:
1. **Filas duplicadas en Gestión de Usuarios**: En el grid de usuarios aparecían filas duplicadas para un mismo `nombreusuario` debido a inserciones históricas con identificadores GUID pseudoaleatorios (`randomblob(16)`), duplicación de semillas y fusiones de réplicas en OneDrive vía CR-SQLite.
2. **Pérdida del cambio de subunidad al reiniciar Excel**: Cuando un usuario cambiaba su subunidad territorial en `FormUsuarios` y guardaba, el cambio no persistía tras reiniciar Excel porque el formulario ejecutaba sentencias SQL directas sobre la base de datos temporal sin vaciar el WAL ni re-encriptar/persistir el vault `.db.enc` en disco. Adicionalmente, el fallback en `config` ejecutaba sentencias `UPDATE config` sin filtrar por `WHERE id = 1`, arriesgando mutaciones masivas.

## Changes

1. **`ConversorPDF.XLL/Academic/Infrastructure/Context/AcademicDbSession.cs`**:
   - Incorpora el método `Persist()` que ejecuta un checkpoint de truncado de WAL (`PRAGMA wal_checkpoint(TRUNCATE);`), extrae los bytes desencriptados de la base temporal con reintentos y solicita a `EncryptedSQLiteContext` la re-encriptación y persistencia atómica en disco (`ReEncryptAndPersist`).

2. **`ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/ConfigRepository.cs`** e **`IConfigRepository.cs`**:
   - Añade el método `UpdateSubunidad(string subunidad)` que actualiza exclusivamente el registro canónico de configuración (`UPDATE config SET subunidad = @Subunidad WHERE id = 1;`).
   - Añade sobrecarga de constructor `ConfigRepository(IDbConnection db)` para soportar invocación directa con conexiones activas.

3. **`ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`**:
   - Inyección de dependencias de `AcademicDbSession` en el constructor.
   - Eliminación de todas las sentencias SQL crudas en el code-behind (`db.Execute("UPDATE config...")`).
   - Uso de `IConfigRepository` para mutar la subunidad del usuario actual.
   - Invocación de `_session.Persist()` en los eventos de guardado y eliminación para asegurar que todos los cambios se graben en el vault encriptado `.db.enc`.
   - Normalización de identificadores a formato determinista `usr_<nombreusuario>`.

4. **`ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`**:
   - `GetUsuarios()`: Deduplica las filas leídas mediante `ROW_NUMBER() OVER (PARTITION BY LOWER(TRIM(nombreusuario)) ORDER BY ...)` filtrando `rn = 1`, conservando siempre la fila con datos más ricos (firma > perfil != 'Academica' > updated_at DESC > id ASC).
   - `InsertUsuario()`: Asigna identificadores deterministas `usr_<lower(trim(nombreusuario))>`.

5. **`ConversorPDF.XLL/Academic/Infrastructure/Persistence/Sql/usuarios.sql`**:
   - Semillas actualizadas con identificadores deterministas (`usr_usuario`, `usr_luis.bechara`, `usr_henry.castillo`, etc.).

6. **`ConversorPDF.XLL/Core/Sync/SubunidadResolver.cs`**:
   - Ordenamiento determinista `ORDER BY activo DESC, updated_at DESC, id ASC LIMIT 1`.
   - Logging estructurado con identificadores y subunidades candidatas cuando existen múltiples registros.
   - Coherencia en el fallback de la tabla `config`.

7. **`ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV79_NormalizeUsuariosIds.cs`**:
   - Migración V79 idempotente (aditiva).
   - Crea copia de seguridad `.bak` antes de mutar.
   - Deduplica filas conservando el registro más rico.
   - Normaliza identificadores legacy a `usr_<nombreusuario>` usando `DELETE` + `INSERT` para preservar tombstones y relojes CR-SQLite.
   - Ejecuta `PRAGMA integrity_check`.

8. **`ConversorPDF.XLL/Core/Sync/SyncInitializer.cs`**:
   - Incorpora el método idempotente `DeduplicateUsuarios(IDbConnection)`.
   - Se ejecuta en `EnableAllCrdts` inmediatamente después de la comprobación de esquemas y antes de regenerar o autosanar triggers CRDT, previniendo que pulls remotos reintroduzcan duplicados.

9. **`ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs`**:
   - Formalizado con Ruling R-20261027-T6-01: la conmutación entre vaults es gestionada de forma natural por `SyncContext.ResolveSubunidadWithDialog` en el ciclo de arranque a partir del vault persistido.

10. **`Ribbon/RibbonController.cs`**:
    - Actualizado el call-site de instanciación de `FormUsuarios` inyectando `session`.

## How to use
1. En la pestaña de ribbon, abrir **Gestión de Usuarios**.
2. La grilla mostrará una única fila por usuario sin registros repetidos.
3. Al seleccionar un usuario y cambiar su subunidad territorial en el combo:
   - Al pulsar **Guardar**, el sistema actualiza la fila, la configuración global y persiste de inmediato el vault en disco.
   - El mensaje confirmará la actualización y advertirá de reiniciar Excel si se modificó el usuario actual.
4. Al cerrar y reabrir Excel:
   - `SyncContext` detecta la nueva subunidad del usuario y conmuta limpiamente al vault de la subunidad seleccionada.

## Test coverage
- `ConversorPDF.Tests.Integration.Usuarios.UsuariosPersistenciaTests`: Verificación E2E de ciclo de vida del vault, persistencia de `config.subunidad` y reapertura del contexto encriptado en disco.
- `ConversorPDF.Tests.Integration.Usuarios.UsuariosDuplicadosTests`:
  - `GetUsuarios_WhenDuplicateUsersExist_ShouldReturnSingleUserWithRichestData`: Deduplicación en lectura conservando firma y rol.
  - `InsertUsuario_WhenDeterministicId_ShouldUseUsrPrefixAndLowerUsername`: Generación de identidad determinista.
  - `MigrationV79_WhenExecuted_ShouldDeduplicateAndNormalizeIds`: Migración V79 sobre base de datos migrada.
  - `DeduplicateUsuarios_WhenPullReintroducesDuplicate_ShouldRetainRichestRow`: Deduplicación post-merge de sync.
- `ConversorPDF.Tests.Unit.Usuarios.SubunidadResolverTests`: Resolución determinista con registros múltiples y nulos en `updated_at`.
- **Suite completa**: 946 pruebas superadas (0 fallos).
