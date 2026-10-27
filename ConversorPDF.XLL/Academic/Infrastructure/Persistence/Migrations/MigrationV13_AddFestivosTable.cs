using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración v13 para crear la tabla festivos en bases de datos existentes.
    /// </summary>
    public class MigrationV13_AddFestivosTable : IMigration
    {
        public int Version => 13;
        public string Name => "Crear tabla festivos para el control de días festivos por período académico";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            connection.Execute(@"
                CREATE TABLE IF NOT EXISTS festivos (
                    idfestivo INTEGER PRIMARY KEY AUTOINCREMENT,
                    idperiodo INTEGER,
                    fechafestivoinicio TEXT,
                    fechafestivofinal TEXT,
                    detallefestivo TEXT,
                    FOREIGN KEY(idperiodo) REFERENCES periodos(idperiodo)
                );
            ", transaction: transaction);
        }
    }
}
