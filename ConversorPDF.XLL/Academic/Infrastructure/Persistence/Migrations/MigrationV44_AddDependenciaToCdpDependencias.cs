using System;
using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Fase 11: Migración para agregar la columna dependencia a instalaciones existentes pre-V28.
    /// Resuelve el SQLiteException al intentar guardar en tablas donde la columna no existía.
    /// </summary>
    public class MigrationV44_AddDependenciaToCdpDependencias : IMigration
    {
        public int Version => 44;
        public string Name => "Add dependencia column to cdpDependencias";

        public void Up(IDbConnection db, IDbTransaction transaction)
        {
            var cols = db.Query<string>(
                "SELECT name FROM pragma_table_info('cdpDependencias');",
                transaction: transaction);

            if (!cols.Any(c => string.Equals(c, "dependencia", StringComparison.OrdinalIgnoreCase)))
            {
                db.Execute(
                    "ALTER TABLE cdpDependencias ADD COLUMN dependencia TEXT DEFAULT NULL;",
                    transaction: transaction);
            }
        }
    }
}
