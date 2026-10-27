using System;
using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// FASE 26: Agrega la columna subunidad a cdpDependencias en instalaciones existentes.
    /// AcademicSchema.sql se actualizó para instalaciones nuevas, pero las existentes
    /// necesitan esta migración para soportar CRDT local.
    /// </summary>
    public class MigrationV34_AddSubunidadToCdpDependencias : IMigration
    {
        public int Version => 34;
        public string Name => "Add subunidad column to cdpDependencias";

        public void Up(IDbConnection db, IDbTransaction transaction)
        {
            // SQLite no soporta IF NOT EXISTS en ALTER TABLE, así que verificamos con PRAGMA
            var cols = db.Query<string>(
                "SELECT name FROM pragma_table_info('cdpDependencias');",
                transaction: transaction);

            if (!cols.Any(c => string.Equals(c, "subunidad", StringComparison.OrdinalIgnoreCase)))
            {
                db.Execute(
                    "ALTER TABLE cdpDependencias ADD COLUMN subunidad TEXT;",
                    transaction: transaction);
            }
        }
    }
}
