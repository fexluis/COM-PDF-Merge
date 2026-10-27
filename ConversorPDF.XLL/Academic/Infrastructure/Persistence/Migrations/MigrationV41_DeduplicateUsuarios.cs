using System;
using System.Data;
using System.Data.SQLite;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Sync;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// MIGRACIÓN V41: Deduplicación de Usuarios y generación de Tombstones.
    ///
    /// CONTEXTO:
    ///   Debido a un comportamiento agresivo del MIN(id) en SyncInitializer 
    ///   combinado con la sincronización multi-dispositivo de la semilla (V29),
    ///   se crearon y esparcieron registros duplicados de usuarios que carecían
    ///   de la información ingresada por los usuarios.
    /// 
    /// SOLUCIÓN:
    ///   - Activar temporalmente CR-SQLite en la conexión de migración.
    ///   - Usar ROW_NUMBER() para identificar el registro "más rico en datos" 
    ///     y eliminar (DELETE) los duplicados con la misma cuenta de usuario.
    ///   - Al estar CR-SQLite activo, los deletes se convierten automáticamente
    ///     en Tombstones (en usuarios__crsql_clock) que viajarán a OneDrive
    ///     limpiando la base de datos de los demás computadores.
    /// </summary>
    public class MigrationV41_DeduplicateUsuarios : IMigration
    {
        public int Version => 41;
        public string Name => "Deduplicate Usuarios and Sync Deletions";

        public void Up(IDbConnection db, IDbTransaction transaction)
        {
            LoggerHelper.Info("[MigrationV41] Iniciando limpieza inteligente de usuarios duplicados...");

            // Paso 1: Habilitar CR-SQLite si no estaba activo en esta conexión
            // Esto es crucial para que los DELETE generen tombstones
            if (db is SQLiteConnection sqliteConn)
            {
                bool loaded = CrSqliteLoader.EnableOnConnection(sqliteConn);
                if (!loaded)
                {
                    LoggerHelper.Warning("[MigrationV41] No se pudo cargar CR-SQLite. Los borrados no se propagarán correctamente.");
                }
                else
                {
                    LoggerHelper.Debug("[MigrationV41] CR-SQLite habilitado exitosamente para la migración.");
                }
            }

            // Paso 2: Deduplicar usando ROW_NUMBER()
            // Prioridad: 
            // 1. Tiene firma
            // 2. Tiene perfil diferente al por defecto ("Academica")
            // 3. Más reciente (updated_at)
            // 4. Último recurso: UUID alfabéticamente
            string sqlDeleteDuplicates = @"
                DELETE FROM usuarios 
                WHERE id NOT IN (
                    SELECT id FROM (
                        SELECT id,
                               ROW_NUMBER() OVER (
                                   PARTITION BY LOWER(nombreusuario) 
                                   ORDER BY 
                                       CASE WHEN firma IS NOT NULL AND firma != '' THEN 1 ELSE 0 END DESC,
                                       CASE WHEN perfil IS NOT NULL AND perfil != 'Academica' THEN 1 ELSE 0 END DESC,
                                       updated_at DESC,
                                       id ASC
                               ) as rn
                        FROM usuarios
                    )
                    WHERE rn = 1
                );
            ";

            int rowsDeleted = db.Execute(sqlDeleteDuplicates, transaction: transaction);
            
            LoggerHelper.Info($"[MigrationV41] Deduplicación completada. Se eliminaron {rowsDeleted} usuarios duplicados (Tombstones generados si CRDT activo).");
        }
    }
}
