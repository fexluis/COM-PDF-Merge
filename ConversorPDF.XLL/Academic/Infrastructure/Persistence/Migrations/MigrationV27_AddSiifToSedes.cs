using ConversorPDF.XLL.Core.Helpers;
using Dapper;
using System.Data;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV27_AddSiifToSedes : IMigration
    {
        public int Version => 27;
        public string Name => "Add Siif columns To Sedes";

        public void Up(IDbConnection db, IDbTransaction transaction)
        {
            if (db.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='sedes';", transaction: transaction) == 0)
                return;

            var columns = db.Query("PRAGMA table_info(sedes)").ToList();
            bool hasDepartamentoSiif = columns.Any(c => string.Equals(c.name?.ToString(), "departamentoSiif", System.StringComparison.OrdinalIgnoreCase));
            bool hasCiudadSiif = columns.Any(c => string.Equals(c.name?.ToString(), "ciudadSiif", System.StringComparison.OrdinalIgnoreCase));

            if (!hasDepartamentoSiif)
            {
                db.Execute("ALTER TABLE sedes ADD COLUMN departamentoSiif TEXT;", transaction: transaction);
                LoggerHelper.Info("[MigrationV27] Columna 'departamentoSiif' añadida.");
            }

            if (!hasCiudadSiif)
            {
                db.Execute("ALTER TABLE sedes ADD COLUMN ciudadSiif TEXT;", transaction: transaction);
                LoggerHelper.Info("[MigrationV27] Columna 'ciudadSiif' añadida.");
            }
        }
    }
}
