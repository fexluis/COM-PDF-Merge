using System;
using System.Linq;
using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV31_AddIdPeriodoToConfig : IMigration
    {
        public int Version => 31;
        public string Name => "AddIdPeriodoToConfig";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // Verificamos si la tabla existe primero
            int tableExists = connection.ExecuteScalar<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='config';",
                transaction: transaction);

            if (tableExists == 0) return;

            // Verificamos si la columna existe antes de intentar agregarla
            var columns = connection.Query("PRAGMA table_info(config)", transaction: transaction).Select(r => (string)((dynamic)r).name).ToList();
            if (!columns.Contains("id_periodo", StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE config ADD COLUMN id_periodo INTEGER DEFAULT 0;", transaction: transaction);
            }
        }
    }
}
