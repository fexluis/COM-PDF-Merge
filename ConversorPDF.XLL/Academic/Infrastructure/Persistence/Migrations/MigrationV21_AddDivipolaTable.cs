using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V21: Crea la tabla divipola si no existe.
    /// Permite a instalaciones existentes (pre-v21) obtener el esquema de División Político-Administrativa
    /// sin necesidad de recrear la base de datos desde cero.
    /// </summary>
    public class MigrationV21_AddDivipolaTable : IMigration
    {
        public int Version => 21;
        public string Name => "Añade tabla divipola para División Político-Administrativa";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // IDEMPOTENCIA: Verificar si la tabla divipola ya existe y tiene datos
            // antes de ejecutar el SQL (que contiene cientos de INSERTs sin OR IGNORE).
            int count = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='divipola';",
                transaction: transaction);

            if (count == 0)
            {
                string sql = Core.Helpers.SqliteHelper.GetEmbeddedSqlResource("Academic_divipola.sql");
                connection.Execute(sql, transaction: transaction);
            }
        }
    }
}
