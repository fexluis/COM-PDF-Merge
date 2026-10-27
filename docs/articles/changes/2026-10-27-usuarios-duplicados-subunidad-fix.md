# Correccion posterior: provision de usuario al conmutar de vault

**Fecha**: 2026-10-27 · **Origen**: auditoria independiente del plan `2026-10-27-usuarios-duplicados-subunidad-plan.md`
**Ruling vigente**: `R-20261027-T6-02` (supersede a `R-20261027-T6-01`)

## Proposito

La verificacion independiente detecto que el cambio de subunidad **seguia perdiendose al conmutar de vault**:
`SyncContext.ResolveSubunidadWithDialog` cambiaba de `.db.enc` al arrancar sin provisionar la fila del usuario en el
vault destino. Si el vault destino ya existia, su fila semilla conservaba la subunidad anterior y el arranque volvia
al vault de origen (sintoma: "guardo la subunidad, cierro Excel y al abrir no persiste").

## Cambios

1. **`ConversorPDF.XLL/Core/Sync/UsuarioVaultProvisioner.cs`** (nuevo)
   - Elimina duplicados legacy de la misma cuenta (sobrevive la identidad determinista `usr_<nombreusuario>`).
   - Upsert idempotente con la subunidad resuelta; no pisa el nombre si llega vacio.
2. **`ConversorPDF.XLL/Core/Sync/SyncContext.cs`**
   - Provisiona la fila del usuario en el vault destino tanto al registrarse como al **conmutar** (`vaultChanged`).
   - Antes solo insertaba en el camino de registro y con `id` aleatorio (`INSERT OR IGNORE` sin `id`).
3. **`AcademicRepository.UpdateUsuario`**
   - Sella `updated_at = datetime('now')`; el criterio "mas reciente" del dedupe era inerte porque nunca cambiaba.
4. **`AcademicRepository.GetUsuarios`** y **`SyncInitializer.DeduplicateUsuarios`**
   - Priorizan `substr(id, 1, 4) = 'usr_'` para que un duplicado legacy con `updated_at` mas nuevo no resucite un id GUID.

## Tests

| Archivo | Casos |
|---|---|
| `ConversorPDF.Tests/Integration/Usuarios/UsuarioVaultProvisionerTests.cs` | semilla con subunidad vieja, fila con id legacy, alta nueva con id determinista, idempotencia |
| `ConversorPDF.Tests/Integration/Usuarios/UsuariosRegresionTests.cs` | `updated_at` se actualiza al guardar; `GetUsuarios` prefiere el id determinista |

Resultado: `verify.ps1 -Mode Fast -Filter "FullyQualifiedName~Usuarios"` -> **13/13 PASS**; `Full` sin regresiones.

## Cobertura declarada

El *provisioner* esta cubierto por tests. La invocacion desde `SyncContext` (metodo modal que abre
`FormRegistroSubunidad`) se valida por revision de codigo, no por test E2E automatizado.

## Nota sobre el documento anterior

`docs/articles/changes/2026-10-27-usuarios-duplicados-subunidad.md` afirmaba resolver los dos defectos. Con esta
correccion, el segundo (persistencia del cambio de subunidad a traves de la conmutacion de vault) queda cubierto.
