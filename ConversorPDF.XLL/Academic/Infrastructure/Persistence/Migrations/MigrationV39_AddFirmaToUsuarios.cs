using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV39_AddFirmaToUsuarios : IMigration
    {
        public int Version => 39;
        public string Name => "AddFirmaToUsuarios";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            int exists = connection.ExecuteScalar<int>(@"
                SELECT COUNT(1) FROM pragma_table_info('usuarios')
                WHERE name = 'firma';", transaction: transaction);

            if (exists == 0)
            {
                connection.Execute(
                    "ALTER TABLE usuarios ADD COLUMN firma TEXT DEFAULT NULL;",
                    transaction: transaction);
            }
        }
    }
}
