using System;
using System.Data;
using System.Data.SQLite;
using System.IO;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Sync;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V79: Deduplicación y normalización determinista de identidades de usuarios (usr_{nombreusuario}).
    /// </summary>
    public sealed class MigrationV79_NormalizeUsuariosIds : IMigration, IMigrationTolerante
    {
        public int Version => 79;
        public string Name => "NormalizeUsuariosIds";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction = null)
        {
            LoggerHelper.Info("[V79] Iniciando migración V79: normalización de IDs y deduplicación de usuarios...");

            // 1. Respaldo de seguridad previo a mutaciones
            CrearRespaldoSeguridad(connection);

            // 2. Habilitar CR-SQLite si es SQLiteConnection para que los borrados emitan tombstones válidos
            if (connection is SQLiteConnection sqliteConn)
            {
                bool crLoaded = CrSqliteLoader.EnableOnConnection(sqliteConn);
                if (crLoaded)
                {
                    LoggerHelper.Debug("[V79] CR-SQLite habilitado en la conexión para generación de tombstones.");
                }
                else
                {
                    LoggerHelper.Warning("[V79] CR-SQLite no disponible en la conexión actual.");
                }
            }

            // Verificar si la tabla usuarios existe
            int tablaExiste = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='usuarios';",
                transaction: transaction);

            if (tablaExiste == 0)
            {
                LoggerHelper.Info("[V79] Tabla usuarios no encontrada. Omitiendo migración V79.");
                return;
            }

            // 3. Deduplicar conservando la fila más rica:
            //    Prioridad: 1. Tiene firma, 2. Perfil no genérico ('Academica'), 3. updated_at DESC, 4. id ASC
            string sqlDeduplicate = @"
                DELETE FROM usuarios 
                WHERE id NOT IN (
                    SELECT id FROM (
                        SELECT id,
                               ROW_NUMBER() OVER (
                                   PARTITION BY LOWER(TRIM(nombreusuario)) 
                                   ORDER BY 
                                       CASE WHEN firma IS NOT NULL AND firma != '' THEN 1 ELSE 0 END DESC,
                                       CASE WHEN perfil IS NOT NULL AND perfil != 'Academica' THEN 1 ELSE 0 END DESC,
                                       updated_at DESC,
                                       id ASC
                               ) as rn
                        FROM usuarios
                    )
                    WHERE rn = 1
                );";

            int eliminados = connection.Execute(sqlDeduplicate, transaction: transaction);
            LoggerHelper.Info($"[V79] Deduplicación de usuarios completada. Se eliminaron {eliminados} filas duplicadas.");

            // 4. Reescribir IDs legacy a identificadores deterministas usr_<lower(trim(nombreusuario))>
            //    usando DELETE + INSERT para preservar coherencia de relojes CRDT y tombstones.
            string sqlSelectLegacy = @"
                SELECT id, nombre, nombreusuario, subunidad, perfil, activo, firma, created_at, updated_at
                FROM usuarios
                WHERE id != ('usr_' || LOWER(TRIM(nombreusuario)));";

            var rowsLegacy = connection.Query(sqlSelectLegacy, transaction: transaction);
            int normalizados = 0;

            foreach (var row in rowsLegacy)
            {
                string oldId = (string)row.id;
                string rawUser = (string)row.nombreusuario;
                string newId = "usr_" + (rawUser ?? string.Empty).Trim().ToLowerInvariant();

                if (string.Equals(oldId, newId, StringComparison.Ordinal))
                {
                    continue;
                }

                connection.Execute("DELETE FROM usuarios WHERE id = @OldId;", new { OldId = oldId }, transaction: transaction);
                connection.Execute("DELETE FROM usuarios WHERE id = @NewId;", new { NewId = newId }, transaction: transaction);

                connection.Execute(@"
                    INSERT INTO usuarios (id, nombre, nombreusuario, subunidad, perfil, activo, firma, created_at, updated_at)
                    VALUES (@id, @nombre, @nombreusuario, @subunidad, @perfil, @activo, @firma, @created_at, @updated_at);",
                    new
                    {
                        id = newId,
                        nombre = (string)row.nombre,
                        nombreusuario = rawUser,
                        subunidad = (string)row.subunidad,
                        perfil = (string)row.perfil,
                        activo = row.activo,
                        firma = (string)row.firma,
                        created_at = (string)row.created_at,
                        updated_at = (string)row.updated_at
                    }, transaction: transaction);

                normalizados++;
            }

            LoggerHelper.Info($"[V79] Normalización completada. Se normalizaron {normalizados} identidades a formato determinista.");

            // 5. Ejecutar PRAGMA integrity_check
            try
            {
                string integrity = connection.QueryFirstOrDefault<string>("PRAGMA integrity_check;", transaction: transaction);
                LoggerHelper.Info($"[V79] Verificación de integridad completada. Resultado: {integrity}");
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V79] Error ejecutando PRAGMA integrity_check: {ex.Message}");
            }
        }

        private static void CrearRespaldoSeguridad(IDbConnection connection)
        {
            try
            {
                if (connection is SQLiteConnection sqliteConn && !string.IsNullOrEmpty(sqliteConn.FileName))
                {
                    if (File.Exists(sqliteConn.FileName))
                    {
                        string bakTemp = sqliteConn.FileName + ".bak";
                        File.Copy(sqliteConn.FileName, bakTemp, true);
                        LoggerHelper.Info($"[V79] Copia de seguridad temporal creada: {bakTemp}");
                    }
                }

                string vaultPath = SqliteHelper.GetVaultPath();
                if (!string.IsNullOrEmpty(vaultPath) && File.Exists(vaultPath))
                {
                    string bakVault = vaultPath + ".bak";
                    File.Copy(vaultPath, bakVault, true);
                    LoggerHelper.Info($"[V79] Copia de seguridad de vault creada: {bakVault}");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V79] No se pudo crear copia de seguridad previa a migración: {ex.Message}");
            }
        }
    }
}
