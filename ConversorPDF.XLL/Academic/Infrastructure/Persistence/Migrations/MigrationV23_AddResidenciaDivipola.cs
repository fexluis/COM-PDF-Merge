using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV23_AddResidenciaDivipola : IMigration
    {
        public int Version => 23;

        public string Name => "Add ResidenciaDivipola To Terceros";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='terceros';", transaction: transaction) == 0)
                return;

            // 1. Verificar si la columna ya existe (idempotencia)
            var columns = connection.Query("PRAGMA table_info(terceros)").ToList();
            bool hasColumn = columns.Any(c => string.Equals(c.name?.ToString(), "id_residencia_divipola", System.StringComparison.OrdinalIgnoreCase));

            if (!hasColumn)
            {
                // Alterar tabla para agregar la nueva columna
                connection.Execute("ALTER TABLE terceros ADD COLUMN id_residencia_divipola INTEGER NULL;", transaction: transaction);
            }

            // 2. Mapear datos historicos (se debe ejecutar incluso si la columna ya existia para db nuevas)
            string updateSql = @"
                UPDATE terceros SET id_residencia_divipola = 362 WHERE UPPER(residencia) = 'POPAYÁN';
                UPDATE terceros SET id_residencia_divipola = 1006 WHERE UPPER(residencia) = 'CALI';
                UPDATE terceros SET id_residencia_divipola = 605 WHERE UPPER(residencia) = 'NEIVA';
                UPDATE terceros SET id_residencia_divipola = 716 WHERE UPPER(residencia) = 'PASTO';
                UPDATE terceros SET id_residencia_divipola = 1041 WHERE UPPER(residencia) = 'TULUÁ';
                UPDATE terceros SET id_residencia_divipola = 319 WHERE UPPER(residencia) = 'MANIZALES';
                UPDATE terceros SET id_residencia_divipola = 959 WHERE UPPER(residencia) = 'IBAGUE';
                UPDATE terceros SET id_residencia_divipola = 149 WHERE UPPER(residencia) = 'BOGOTÁ';
                UPDATE terceros SET id_residencia_divipola = 150 WHERE UPPER(residencia) = 'CARTAGENA';
                UPDATE terceros SET id_residencia_divipola = 832 WHERE UPPER(residencia) = 'PEREIRA';
                UPDATE terceros SET id_residencia_divipola = 429 WHERE UPPER(residencia) = 'MONTERIA';
                UPDATE terceros SET id_residencia_divipola = 933 WHERE UPPER(residencia) = 'SINCELEJO';
                UPDATE terceros SET id_residencia_divipola = 192 WHERE UPPER(residencia) = 'TURBACO';
                UPDATE terceros SET id_residencia_divipola = 126 WHERE UPPER(residencia) = 'BARRANQUILLA';
            ";
            connection.Execute(updateSql, transaction: transaction);
        }
    }
}
