# Movimiento de Usuarios entre Vaults al Cambiar de Subunidad (MUV)

## Purpose
Este cambio implementa el movimiento físico y atómico de usuarios entre vaults territoriales (`AcademicData_<SUBUNIDAD>.db.enc`) cuando se modifica su subunidad en Gestión de Usuarios (`FormUsuarios`), saldando la deuda técnica pendiente y garantizando:
1. **Existencia física en el vault destino**: El usuario editado (sea el usuario actual o un tercero) queda inmediatamente registrado en `AcademicData_<DESTINO>.db.enc` preservando su perfil, estado activo y firma digital.
2. **Preservación CRDT sin tombstones destructivos (R-MUV-01)**: El movimiento se ejecuta con `UPDATE` en el origen y `UPSERT` en el destino, prohibiendo terminantemente sentencias `DELETE FROM usuarios` que generarían tombstones propagables y borrarían la fila en réplicas de OneDrive.
3. **Coherencia de configuración del dispositivo (R-MUV-02)**: La tabla `config.subunidad` (id=1) representa la configuración física del dispositivo y **solo** se actualiza cuando el usuario modificado coincide con el usuario de Windows actual (`Environment.UserName`).
4. **Recarga de interfaz en vivo (D8)**: Si `FormConfig` se encuentra abierto simultáneamente, su combo de subunidad se refresca de inmediato mediante `RecargarConfiguracion()`.

## Changes

1. **`ConversorPDF.XLL/Academic/Application/Interfaces/IUsuarioVaultMover.cs`**:
   - Puerto de infraestructura desacoplado para ejecutar el movimiento físico hacia vaults destino.

2. **`ConversorPDF.XLL/Academic/Application/Interfaces/IMoverUsuarioSubunidadService.cs`**:
   - Contrato enriquecido: `Result<MovimientoUsuarioResultado> Mover(Usuario usuario, string subunidadDestino, string usuarioActual)`.

3. **`ConversorPDF.XLL/Academic/Application/DTOs/`**:
   - `MovimientoVaultResultado`: Detalle técnico de la operación de persistencia física en el vault destino (`SubunidadDestino`, `RutaVaultDestino`, `Insertado`, `Actualizado`).
   - `MovimientoUsuarioResultado`: Resultado de negocio para la capa UI (`UsuarioId`, `SubunidadDestino`, `ConfigActualizada`, `RequiereReinicio`).

4. **`ConversorPDF.XLL/Academic/Application/Services/MoverUsuarioSubunidadService.cs`**:
   - Orquestador de la lógica de aplicación sin dependencias de infraestructura ni UI. Aplica guardas de dominio, delega en `IUsuarioVaultMover` y encapsula la regla de actualización de `IConfigRepository` exclusivamente para el usuario actual.

5. **`ConversorPDF.XLL/Academic/Infrastructure/Persistence/UsuarioVaultMover.cs`**:
   - Implementación de `IUsuarioVaultMover`.
   - `UpsertEnVault`: Sentencia `INSERT ... ON CONFLICT(id) DO UPDATE` con normalización determinista `usr_<nombreusuario>`, preservando perfil, activo, firma y actualizando `updated_at`.
   - Persistencia segura: Extrae los bytes desencriptados del archivo temporal mediante `EncryptedSQLiteContext.LeerBytesConRetrySync`, cifra y persiste con `SqliteHelper.UpdateVaultDirectly` y marca los cambios persistidos.
   - Constructor `internal` con delegados inyectables para soportar pruebas automatizadas sin tocar vaults en producción.

6. **`ConversorPDF.XLL/Academic/UI/Forms/FormConfig.cs`**:
   - Añade el método público `RecargarConfiguracion()` para recargar los datos y combos en caliente si el formulario está instanciado.

7. **`ConversorPDF.XLL/Academic/UI/Forms/FormUsuarios.cs`**:
   - En `btnGuardar`, ante un cambio de subunidad se compone e invoca `MoverUsuarioSubunidadService`.
   - Persistencia en orden estricto: primero se persiste el vault destino, luego se persiste el vault de origen mediante `_session.Persist()`.
   - Si se actualizó la configuración del usuario actual, busca instancias abiertas de `FormConfig` y solicita su recarga.

8. **`ConversorPDF.XLL/ConversorPDF.XLL.csproj`**:
   - Registro de los nuevos contratos, DTOs e infraestructura en el proyecto.

## How to use
1. En Excel, pestaña **Conversor PDF**, abrir **Gestión de Usuarios**.
2. Seleccionar cualquier usuario (propio o de un tercero) y cambiar el valor de **Subunidad**.
3. Hacer clic en **Guardar**:
   - Si se editó a un **tercero**: el usuario se inserta/actualiza en el vault destino de esa subunidad (`AcademicData_<DESTINO>.db.enc`) y la grilla de origen se persiste. La configuración del dispositivo no cambia.
   - Si se editó al **usuario actual**: además del movimiento de vault, se actualiza `config.subunidad`, `FormConfig` (si está abierto) se actualiza de inmediato y el sistema solicita el reinicio de Excel para conmutar al nuevo vault maestro.

## Test coverage
- **Unit (`ConversorPDF.Tests/Unit/Usuarios/MoverUsuarioSubunidadServiceTests.cs`)**:
  - `Mover_WhenUsuarioNoEsElActual_ShouldNotTouchConfig`
  - `Mover_WhenUsuarioEsElActual_ShouldUpdateConfigAndRequireRestart`
  - `Mover_WhenVaultMoverFalla_ShouldPropagateFailureAndNotTouchConfig`
  - `Mover_WhenSubunidadDestinoVacia_ShouldFail` (Theory)
  - `Mover_WhenUsuarioIdVacio_ShouldFail`
- **Integration SQL (`ConversorPDF.Tests/Integration/Usuarios/UsuarioVaultMoverTests.cs`)**:
  - `Mover_WhenUsuarioExisteEnOrigen_ShouldPreservePerfilActivoYFirma`
  - `Mover_WhenSeInvocaDosVeces_ShouldSerIdempotente`
  - `Mover_WhenUsuarioNoExisteEnDestino_ShouldInsertarConIdDeterminista`
- **E2E Vaults Cifrados (`ConversorPDF.Tests/Integration/Usuarios/MoverUsuarioEntreVaultsE2ETests.cs`)**:
  - `Mover_WhenCambiaSubunidad_ShouldPersistirFilaEnVaultDestinoYReabrirla`
  - `Mover_WhenRutaDestinoIgualAOrigen_ShouldSerNoOpSinPerdida`
- **Suite Completa (`scripts/verify.ps1 -Mode Full`)**:
  - **TOTAL: 964 pruebas aprobadas**, 0 errores, 0 regresiones (Baseline: 952 pruebas).
