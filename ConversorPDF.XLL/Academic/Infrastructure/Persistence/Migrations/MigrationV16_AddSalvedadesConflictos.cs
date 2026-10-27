using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV16_AddSalvedadesConflictos : IMigration
    {
        public int Version => 16;
        public string Name => "Agregar tabla salvedades_conflictos";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            connection.Execute(@"
                CREATE TABLE IF NOT EXISTS salvedades_conflictos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    idsesion_a INTEGER NOT NULL,
                    idsesion_b INTEGER NOT NULL,
                    idtercero INTEGER NOT NULL,
                    idperiodo INTEGER,
                    justificacion TEXT,
                    fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(idsesion_a, idsesion_b)
                );
            ", transaction: transaction);

            connection.Execute(@"
                CREATE INDEX IF NOT EXISTS idx_salvedades_periodo_tercero 
                ON salvedades_conflictos(idperiodo, idtercero);
            ", transaction: transaction);
        }
    }
}
