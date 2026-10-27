using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V52: Aislamiento idempotente de aulas demo en la tabla ag_aula (Fase 233).
    /// Asegura que en vaults existentes las aulas de ejemplo sembradas por V49 se marquen con subunidad = 'DEMO'
    /// para evitar su aparición accidental en catálogos y algoritmos de sedes reales.
    /// </summary>
    public class MigrationV52_IsolateDemoAulas : IMigration, IMigrationTolerante
    {
        public int Version => 52;
        public string Name => "IsolateDemoAulas";
        public bool Aditiva => false;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info("[V52] Paso 1: Verificando existencia de tabla ag_aula...");
            if (connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_aula';",
                transaction: transaction) == 0)
            {
                return;
            }

            ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info("[V52] Paso 2: Aislando aulas demo de ag_aula a subunidad 'DEMO'...");
            const string sql = @"
UPDATE ag_aula 
SET subunidad = 'DEMO'
WHERE COALESCE(subunidad, '') = ''
  AND codigo IN ('A-101', 'A-102', 'A-201', 'LAB-1', 'VIR-01')
  AND idperiodo = 0;";

            int afectadas = connection.Execute(sql, transaction: transaction);
            ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info($"[V52] Aulas demo aisladas: {afectadas}. Migración completada.");
        }
    }
}
