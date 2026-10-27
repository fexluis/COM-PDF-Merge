using ConversorPDF.XLL.Core.Helpers;
using Dapper;
using System.Data;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV22_AddActivoToSedes : IMigration
    {
        public int Version => 22;
        public string Name => "Add Activo To Sedes";

        public void Up(IDbConnection db, IDbTransaction transaction)
        {
            if (db.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='sedes';", transaction: transaction) == 0)
                return;

            // Verify if column already exists
            var columns = db.Query("PRAGMA table_info(sedes)").ToList();
            bool hasActivo = columns.Any(c => string.Equals(c.name?.ToString(), "activo", System.StringComparison.OrdinalIgnoreCase));

            if (!hasActivo)
            {
                db.Execute("ALTER TABLE sedes ADD COLUMN activo INTEGER DEFAULT 1;", transaction: transaction);
                LoggerHelper.Info("[MigrationV22] Columna 'activo' añadida a tabla 'sedes'.");
            }
        }
    }
}
