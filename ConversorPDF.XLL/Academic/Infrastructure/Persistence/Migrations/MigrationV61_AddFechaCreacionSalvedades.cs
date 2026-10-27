using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V61: Agrega la columna fecha_creacion en ag_salvedades_conflictos para compatibilidad con SalvedadConflictoRepository (Fase 514).
    /// </summary>
    public class MigrationV61_AddFechaCreacionSalvedades : IMigration, IMigrationTolerante
    {
        public int Version => 61;
        public string Name => "AddFechaCreacionSalvedades";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V61] Verificando columna fecha_creacion en ag_salvedades_conflictos...");

            try
            {
                int existeCol = connection.QueryFirstOrDefault<int>(@"
SELECT COUNT(*) FROM pragma_table_info('ag_salvedades_conflictos') WHERE name='fecha_creacion';", transaction: transaction);

                if (existeCol == 0)
                {
                    connection.Execute(
                        "ALTER TABLE ag_salvedades_conflictos ADD COLUMN fecha_creacion TEXT DEFAULT '';",
                        transaction: transaction);
                    LoggerHelper.Info("[V61] Columna fecha_creacion añadida a ag_salvedades_conflictos.");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V61] Advertencia agregando fecha_creacion a ag_salvedades_conflictos: {ex.Message}");
            }
        }
    }
}
