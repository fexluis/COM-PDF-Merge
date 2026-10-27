using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Security.Principal;
namespace ConversorPDF.XLL.Core.Services
{
    /// <summary>
    /// Servicio que gestiona la identidad del usuario y la recuperación de firmas digitales desde la base de datos (columna firma).
    /// </summary>
    public class UserService : IUserService
    {

        /// <summary> Obtiene el nombre completo de la identidad de Windows (ej: DOMINIO\Usuario). </summary>
        public string GetUsuarioWindows()
        {
            try
            {
                return WindowsIdentity.GetCurrent().Name;
            }
            catch
            {
                // FASE 4: Fallback genérico para WindowsIdentity a Environment.UserName
                return Environment.UserName?.ToUpper() ?? "UNKNOWN_USER";
            }
        }

        /// <summary>
        /// Recupera la firma en Base64 asociada a un nombre de usuario específico.
        /// Retorna directamente la cadena Base64 desde la base de datos (columna firma).
        /// </summary>
        /// <param name="username">Nombre de usuario de red a buscar.</param>
        /// <returns>Cadena Base64 con la imagen de la firma o null si no se encuentra configurada.</returns>
        public string GetFirmaBase64(string username)
        {
            if (string.IsNullOrWhiteSpace(username)) return null; // Fase 27

            return GetFirmaBase64FromDb(username);
        }

        private string GetFirmaBase64FromDb(string username)
        {
            try
            {
                const string sql = "SELECT firma FROM usuarios WHERE nombreusuario = @username COLLATE NOCASE";

                // Nivel 1: Subunidad activa en memoria
                string vaultPathRead1 = null;
                string subunidad = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad;
                if (!string.IsNullOrEmpty(subunidad))
                {
                    vaultPathRead1 = SqliteHelper.GetDbPathForSubunidad(subunidad);
                    string firma = QueryFirmaFromVault(vaultPathRead1, username, sql);
                    if (!string.IsNullOrEmpty(firma)) return firma;
                }

                // Nivel 2: Vault activo persistido en .active_vault
                // Fase 19: Deduplicar contra Nivel 1
                string activeVaultPath = ConversorPDF.XLL.Core.Sync.SyncContext.GetActiveVaultPath();
                if (!string.IsNullOrEmpty(activeVaultPath) &&
                    !string.Equals(vaultPathRead1, activeVaultPath, StringComparison.OrdinalIgnoreCase))
                {
                    string firma = QueryFirmaFromVault(activeVaultPath, username, sql);
                    if (!string.IsNullOrEmpty(firma)) return firma;
                }

                // Nivel 3: DB por defecto (comportamiento original)
                // Fase 21: Deduplicar Nivel 3
                string vaultPathRead3 = SqliteHelper.DbPathEnc;
                bool nivel3EsDuplicado =
                    string.Equals(vaultPathRead1, vaultPathRead3, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(activeVaultPath, vaultPathRead3, StringComparison.OrdinalIgnoreCase);

                if (!nivel3EsDuplicado)
                {
                    // Fase 15: NO usar bloque using: GetContext() retorna el Shared Context; su Dispose destruiría la caché compartida.
                    var ctxDefault = SqliteHelper.GetContext();
                    var connDefault = ctxDefault.Open();
                    return Dapper.SqlMapper.QueryFirstOrDefault<string>(connDefault, sql, new { username });
                }
                return null;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[UserService] Error resolviendo firma para {username}: {ex.Message}");
                return null;
            }
        }

        private string QueryFirmaFromVault(string vaultPath, string username, string sql)
        {
            try
            {
                using (var ctx = SqliteHelper.GetContextForPath(vaultPath))
                {
                    var conn = ctx.Open();
                    ctx.MarkAsReadOnly(); // Fase 28: Revertido al orden original (después de Open)
                    return Dapper.SqlMapper.QueryFirstOrDefault<string>(conn, sql, new { username });
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[UserService] Error consultando firma en vault {vaultPath}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Obtiene los datos de identificación (iniciales y nombre completo) asociados los usuarios hardcodeados.
        /// Si el usuario no pertenece a esta lista en duro, recibe null y sus hojas no serán estampadas con datos de identificación.
        /// </summary>
        public UserIdentification GetIdentification(string username)
        {
            if (string.IsNullOrEmpty(username)) return null;

            string normalizedUser = username.ToUpper();

            if (normalizedUser == "ANTOVILL" || normalizedUser == "ELIZABETH.CORDOBAO")
            {
                return new UserIdentification("MHavillega", "ANTONIO ALEXANDER VILLEGAS");
            }

            if (normalizedUser == "EDUIN.PALACIOSXXXX")
            {
                return new UserIdentification("MHlbechara", "LUIS EDUARDO BECHARA POLANIA");
            }

            return null;
        }

        /// <summary>
        /// Guarda o actualiza la firma en Base64 asociada a un usuario en la base de datos (columna firma).
        /// </summary>
        public void SaveFirmaBase64(string username, string base64)
        {
            if (string.IsNullOrWhiteSpace(username)) return; // FASE 26
            if (string.IsNullOrWhiteSpace(base64)) return; // FASE 5: prevenir sobrescritura con vacío

            try
            {
                const string sql = "UPDATE usuarios SET firma = @firma WHERE nombreusuario = @username COLLATE NOCASE";

                // FASE 8: Resolver paths antes de operar para evitar doble escritura sobre el mismo vault
                string vaultPathNivel1 = null;
                string subunidad = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad;
                if (!string.IsNullOrEmpty(subunidad))
                {
                    vaultPathNivel1 = SqliteHelper.GetDbPathForSubunidad(subunidad);
                }

                string vaultPathNivel2 = ConversorPDF.XLL.Core.Sync.SyncContext.GetActiveVaultPath();

                // FASE 23: Trazabilidad de fallo sistémico
                int nivelesEjecutados = 0;
                int nivelesExitosos = 0;

                // Nivel 1: Subunidad activa
                if (!string.IsNullOrEmpty(vaultPathNivel1))
                {
                    nivelesEjecutados++;
                    if (ExecuteUpdateFirma(vaultPathNivel1, sql, username, base64)) nivelesExitosos++;
                }

                // Nivel 2: Vault activo — solo si es un path distinto al Nivel 1
                if (!string.IsNullOrEmpty(vaultPathNivel2) &&
                    !string.Equals(vaultPathNivel1, vaultPathNivel2, StringComparison.OrdinalIgnoreCase))
                {
                    nivelesEjecutados++;
                    if (ExecuteUpdateFirma(vaultPathNivel2, sql, username, base64)) nivelesExitosos++;
                }

                // FASE 11: Deduplicación Nivel 3 contra Nivel 1 y 2
                string vaultPathDefault = SqliteHelper.DbPathEnc;
                bool nivel3EsDuplicado =
                    string.Equals(vaultPathNivel1, vaultPathDefault, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(vaultPathNivel2, vaultPathDefault, StringComparison.OrdinalIgnoreCase);

                if (!nivel3EsDuplicado)
                {
                    nivelesEjecutados++;
                    // Nivel 3: DB por defecto — usar GetContextForPath para obtener contexto desechable
                    // Fase 16: Pasar vaultPathDefault explícitamente (no usar null)
                    using (var ctxDefault = SqliteHelper.GetContextForPath(vaultPathDefault))
                    {
                        var connDefault = ctxDefault.Open();
                        int filas = Dapper.SqlMapper.Execute(connDefault, sql, new { username, firma = base64 });

                        if (filas == 0)
                        {
                            LoggerHelper.Warning($"[UserService] UPDATE de firma no afectó filas. Usuario '{username}' no encontrado.");
                        }
                        else
                        {
                            // Persistir cambios al vault encriptado
                            // FASE 7: Capturar TempDbPath antes de CloseConnectionResiliently por si el cleanup interno lo borra
                            string tempPathSnapshot = ctxDefault.TempDbPath;
                            string encryptedPathSnapshot = ctxDefault.EncryptedFilePath; // FASE 12
                            ctxDefault.CloseConnectionResiliently();
                            byte[] plainBytes = null;
                            if (!string.IsNullOrEmpty(tempPathSnapshot) && System.IO.File.Exists(tempPathSnapshot))
                            {
                                plainBytes = ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context.EncryptedSQLiteContext.LeerBytesConRetrySync(tempPathSnapshot);
                            }

                            if (plainBytes != null && plainBytes.Length > 0)
                            {
                                SqliteHelper.UpdateVaultDirectly(plainBytes, encryptedPathSnapshot);
                                ctxDefault.MarkChangesPersisted();
                                nivelesExitosos++;
                            }
                            else if (filas > 0)
                            {
                                LoggerHelper.Warning($"[UserService] No se pudo leer TempDbPath para persistir firma en vault por defecto.");
                            }
                        }
                    }
                }
                else
                {
                    LoggerHelper.Info("[UserService] Nivel 3 omitido — vault por defecto ya fue actualizado en Nivel 1 o 2.");
                }

                if (nivelesEjecutados > 0 && nivelesExitosos == 0)
                {
                    throw new InvalidOperationException($"No se pudo persistir la firma en ningún vault para el usuario '{username}'. Revise los logs para más detalles.");
                }

                if (nivelesExitosos > 0)
                {
                    try
                    {
                        ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();
                    }
                    catch (Exception markEx)
                    {
                        // Fase 25: MarkDirty no crítico — no propagar. La firma ya fue guardada.
                        LoggerHelper.Warning($"[UserService] MarkDirty falló tras guardar firma: {markEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[UserService] Error guardando firma para {username}: {ex.Message}", ex);
                throw;
            }
        }

        private bool ExecuteUpdateFirma(string vaultPath, string sql, string username, string base64)
        {
            try
            {
                using (var ctx = SqliteHelper.GetContextForPath(vaultPath))
                {
                    var conn = ctx.Open();
                    int filas = Dapper.SqlMapper.Execute(conn, sql, new { username, firma = base64 });

                    if (filas == 0)
                    {
                        LoggerHelper.Warning($"[UserService] UPDATE de firma no afectó filas en vault {vaultPath}. Usuario '{username}' no encontrado.");
                        return false;
                    }
                    else
                    {
                        // Persistir cambios al vault encriptado
                        // FASE 7: Capturar TempDbPath antes de CloseConnectionResiliently por si el cleanup interno lo borra
                        string tempPathSnapshot = ctx.TempDbPath;
                        ctx.CloseConnectionResiliently();
                        byte[] plainBytes = null;
                        if (!string.IsNullOrEmpty(tempPathSnapshot) && System.IO.File.Exists(tempPathSnapshot))
                        {
                            plainBytes = ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context.EncryptedSQLiteContext.LeerBytesConRetrySync(tempPathSnapshot);
                        }

                        if (plainBytes != null && plainBytes.Length > 0)
                        {
                            SqliteHelper.UpdateVaultDirectly(plainBytes, vaultPath);
                            ctx.MarkChangesPersisted();
                            return true;
                        }
                        else
                        {
                            LoggerHelper.Warning($"[UserService] No se pudo leer TempDbPath para persistir firma en vault {vaultPath}.");
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[UserService] Error actualizando firma en vault {vaultPath}: {ex.Message}");
                return false;
            }
        }
    }
}
