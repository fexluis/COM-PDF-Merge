using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV20_AddTipoFestivo : IMigration
    {
        public int Version => 20;
        public string Name => "Agregar columna tipofestivo a festivos";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            try
            {
                // 1. Agregar columna (sin DEFAULT para evitar error si ya existe)
                connection.Execute("ALTER TABLE festivos ADD COLUMN tipofestivo TEXT;", transaction: transaction);
            }
            catch
            {
                // Columna ya existe (schema base v9+)
            }

            // 2. Normalizar filas legacy: NULL → 'FESTIVO'
            connection.Execute(
                "UPDATE festivos SET tipofestivo = 'FESTIVO' WHERE tipofestivo IS NULL;",
                transaction: transaction);
        }
    }
}
