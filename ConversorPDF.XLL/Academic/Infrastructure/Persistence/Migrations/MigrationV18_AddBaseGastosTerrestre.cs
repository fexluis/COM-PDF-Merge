using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV18_AddBaseGastosTerrestre : IMigration
    {
        public int Version => 18;
        public string Name => "Crear tabla BaseGastosTerrestre";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            connection.Execute(@"
                CREATE TABLE IF NOT EXISTS BaseGastosTerrestre (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Territorial TEXT NOT NULL,
                    CodigoTrayecto TEXT NOT NULL UNIQUE,
                    Trayecto TEXT NOT NULL,
                    Valor REAL
                );
            ", transaction: transaction);
        }
    }
}
