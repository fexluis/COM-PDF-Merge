using System;
using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Fase 21: Migración para agregar la columna dependencia_full a la tabla cdpDependencias en SQLite.
    /// Permite almacenar el nombre largo oficial (ej: "SUBDIRECCIÓN ACADÉMICA - PREGRADO") de la dependencia.
    /// </summary>
    public class MigrationV46_AddDependenciaFullToCdpDependencias : IMigration
    {
        public int Version => 46;
        public string Name => "Add dependencia_full column to cdpDependencias";

        public void Up(IDbConnection db, IDbTransaction transaction)
        {
            int tableExists = db.ExecuteScalar<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='cdpDependencias';",
                transaction: transaction);
            if (tableExists == 0)
            {
                db.Execute(@"
                    CREATE TABLE IF NOT EXISTS cdpDependencias (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        idperiodo INTEGER,
                        cdpnum TEXT,
                        fechacdp TEXT,
                        rubro TEXT,
                        responsable TEXT,
                        cargo TEXT,
                        elaboro TEXT,
                        reviso TEXT,
                        dependencia TEXT DEFAULT NULL,
                        dependencia_full TEXT DEFAULT NULL,
                        activo INTEGER DEFAULT 0,
                        subunidad TEXT
                    );", transaction: transaction);
                return;
            }

            var cols = db.Query("PRAGMA table_info(cdpDependencias)", transaction: transaction)
                         .Select(r => (string)((dynamic)r).name)
                         .ToList();

            if (!cols.Contains("dependencia_full", StringComparer.OrdinalIgnoreCase))
            {
                db.Execute(
                    "ALTER TABLE cdpDependencias ADD COLUMN dependencia_full TEXT DEFAULT NULL;",
                    transaction: transaction);
            }
        }
    }
}
