using System;
using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V50: Agrega las columnas de ventana temporal a la tabla periodos y realiza el backfill de fechas en vaults existentes.
    /// </summary>
    public class MigrationV50_AddPeriodoVentana : IMigration, IMigrationTolerante
    {
        public int Version => 50;
        public string Name => "AddPeriodoVentana";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info("[V50] Paso 1: verificando existencia de tabla periodos...");
            if (connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='periodos';",
                transaction: transaction) == 0) { return; }

            var cols = connection.Query<string>("SELECT name FROM pragma_table_info('periodos');", transaction: transaction)
                                 .Select(x => x.ToLowerInvariant()).ToList();
            ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info($"[V50] Columnas actuales en periodos: {string.Join(", ", cols)}");

            if (!cols.Contains("periodo_inicio"))
            {
                ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info("[V50] Paso 2: ALTER TABLE periodos ADD COLUMN periodo_inicio TEXT");
                try
                {
                    connection.Execute("ALTER TABLE periodos ADD COLUMN periodo_inicio TEXT;", transaction: transaction);
                }
                catch (Exception ex) when (ex.Message.IndexOf("duplicate column", StringComparison.OrdinalIgnoreCase) >= 0) { }
            }

            if (!cols.Contains("periodo_fin"))
            {
                ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info("[V50] Paso 3: ALTER TABLE periodos ADD COLUMN periodo_fin TEXT");
                try
                {
                    connection.Execute("ALTER TABLE periodos ADD COLUMN periodo_fin TEXT;", transaction: transaction);
                }
                catch (Exception ex) when (ex.Message.IndexOf("duplicate column", StringComparison.OrdinalIgnoreCase) >= 0) { }
            }

            // FASE 754: Verificar si realmente se requiere backfill antes de soltar triggers CRDT
            bool requiereBackfill = !cols.Contains("periodo_inicio") || !cols.Contains("periodo_fin");
            if (!requiereBackfill)
            {
                requiereBackfill = connection.QueryFirstOrDefault<int>(
                    @"SELECT COUNT(*) FROM periodos
                      WHERE periodo_inicio IS NULL OR periodo_inicio = ''
                         OR periodo_fin IS NULL OR periodo_fin = '';",
                    transaction: transaction) > 0;
            }

            if (requiereBackfill)
            {
                // FASE 747/754: Soltar triggers CRDT obsoletos ÚNICAMENTE si efectivamente se ejecutará backfill DML
                PeriodoVentanaCrdtRepair.DropStaleCrsqlTriggers(connection, transaction);

                ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info("[V50] Paso 4: Backfill parametrizado de fechas en periodos...");
                // Backfill por parámetros: imposible el mismatch de conteo de VALUES
                var datos = new[]
                {
                    new { Id = 1, Ini = "2026-01-12", Fin = "2026-06-19" },
                    new { Id = 2, Ini = "2026-08-15", Fin = "2026-12-12" },
                    new { Id = 3, Ini = "2027-01-01", Fin = "2027-06-30" }
                };

                foreach (var d in datos)
                {
                    connection.Execute(
                        @"UPDATE periodos
                             SET periodo_inicio = COALESCE(NULLIF(periodo_inicio, ''), @Ini),
                                 periodo_fin    = COALESCE(NULLIF(periodo_fin, ''),    @Fin)
                           WHERE idperiodo = @Id;",
                        d, transaction: transaction);
                }
            }
            else
            {
                ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info("[V50] Columnas y fechas completas: sin cambios ni drop de triggers.");
            }

            ConversorPDF.XLL.Core.Helpers.LoggerHelper.Info("[V50] Migración completada exitosamente.");
        }
    }
}
