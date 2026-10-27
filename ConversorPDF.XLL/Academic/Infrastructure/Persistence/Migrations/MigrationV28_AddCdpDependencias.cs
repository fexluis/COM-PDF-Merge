using Dapper;
using System.Data;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV28_AddCdpDependencias : IMigration
    {
        public int Version => 28;
        public string Name => "Add cdpDependencias table";

        public void Up(IDbConnection db, IDbTransaction transaction)
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
                    activo INTEGER DEFAULT 0
                );
            ", transaction: transaction);

        }
    }
}
