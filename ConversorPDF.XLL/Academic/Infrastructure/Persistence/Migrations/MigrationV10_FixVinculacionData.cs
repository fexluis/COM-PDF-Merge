using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV10_FixVinculacionData : IMigration
    {
        public int Version => 10;
        public string Name => "Corrección de datos semilla — Vinculación ID 120 (idprograma, total_horas)";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction) == 0)
                return;

            connection.Execute(
                @"UPDATE vinculaciones 
                  SET idprograma = 2, total_horas = 32 
                  WHERE id = 120 
                    AND (idprograma != 2 OR total_horas != 32);",
                transaction: transaction);
        }
    }
}
