using System;
using System.Linq;
using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// MIGRACIÓN V45: Migración puente para compatibilidad de esquema (Fase 34).
    /// Agrega la columna 'id_territorial' a la tabla 'config' para bases de datos migradas desde esquemas antiguos.
    /// Recupera la correspondencia del id desde la tabla 'territoriales' basándose en la columna 'territorial' de texto.
    /// </summary>
    public class MigrationV45_AddIdTerritorialToConfig : IMigration
    {
        public int Version => 45;
        public string Name => "AddIdTerritorialToConfig";

        /// <summary>
        /// Aplica la migración asegurando resiliencia frente a bases de datos en distintos estados.
        /// Valida la existencia previa de la tabla antes de alterar columnas.
        /// </summary>
        /// <param name="connection">Conexión SQLite abierta y operativa.</param>
        /// <param name="transaction">Transacción activa provista por el SchemaMigrator.</param>
        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            int tableExists = connection.ExecuteScalar<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='config';",
                transaction: transaction);

            if (tableExists == 0) return;

            var columns = connection.Query("PRAGMA table_info(config)", transaction: transaction)
                .Select(r => (string)((dynamic)r).name).ToList();

            if (!columns.Contains("id_territorial", StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE config ADD COLUMN id_territorial INTEGER DEFAULT 0;", transaction: transaction);
                
                // Intentar mapear territorial TEXT a id_territorial INTEGER si existe
                if (columns.Contains("territorial", StringComparer.OrdinalIgnoreCase))
                {
                    connection.Execute(@"
                        UPDATE config 
                        SET id_territorial = COALESCE(
                            (SELECT id FROM territoriales WHERE territorial = config.territorial),
                            (SELECT id FROM territoriales WHERE subunidad = config.subunidad),
                            0
                        )
                        WHERE id_territorial = 0;", transaction: transaction);
                }
            }
        }
    }
}
