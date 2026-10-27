using System;
using System.Data;
using System.Linq;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V65: Reparación definitiva de ventana temporal en la tabla periodos para resolver el hueco de migración V50/V51.
    /// FASE 745: Desactiva triggers stale de CR-SQLite, añade físicamente periodo_inicio y periodo_fin si faltan,
    /// y ejecuta el backfill canónico preservando fechas existentes (id 1: 2026-01-12/2026-06-19, id 2: 2026-08-11/2026-12-12, id 3: 2027-01-01/2027-06-30).
    /// </summary>
    public class MigrationV65_PeriodoVentanaCrdtRepair : IMigration, IMigrationTolerante
    {
        public int Version => 65;
        public string Name => "PeriodoVentanaCrdtRepair";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V65] Paso 1: Verificando existencia de tabla periodos...");
            int tablaExiste = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='periodos';",
                transaction: transaction);

            if (tablaExiste == 0)
            {
                LoggerHelper.Info("[V65] Tabla periodos no existe. Omitiendo migración.");
                return;
            }

            var cols = connection.Query<string>("SELECT name FROM pragma_table_info('periodos');", transaction: transaction)
                                 .Select(x => x.ToLowerInvariant())
                                 .ToList();

            LoggerHelper.Info($"[V65] Columnas actuales en periodos: {string.Join(", ", cols)}");

            // FASE 752: Determinar si el esquema requiere alteración o backfill
            bool faltaIni = !cols.Contains("periodo_inicio");
            bool faltaFin = !cols.Contains("periodo_fin");
            bool requiereBackfill = faltaIni || faltaFin;

            // Paso 2: Agregar columnas faltantes
            if (faltaIni)
            {
                LoggerHelper.Info("[V65] Paso 2a: ALTER TABLE periodos ADD COLUMN periodo_inicio TEXT");
                try
                {
                    connection.Execute("ALTER TABLE periodos ADD COLUMN periodo_inicio TEXT;", transaction: transaction);
                }
                catch (Exception ex) when (ex.Message.IndexOf("duplicate column", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    LoggerHelper.Warning("[V65] Columna periodo_inicio ya existía.");
                }
            }

            if (faltaFin)
            {
                LoggerHelper.Info("[V65] Paso 2b: ALTER TABLE periodos ADD COLUMN periodo_fin TEXT");
                try
                {
                    connection.Execute("ALTER TABLE periodos ADD COLUMN periodo_fin TEXT;", transaction: transaction);
                }
                catch (Exception ex) when (ex.Message.IndexOf("duplicate column", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    LoggerHelper.Warning("[V65] Columna periodo_fin ya existía.");
                }
            }

            // FASE 752: Si las columnas ya existían, verificar si hay fechas incompletas o nulas
            if (!requiereBackfill)
            {
                requiereBackfill = connection.QueryFirstOrDefault<int>(
                    @"SELECT COUNT(*) FROM periodos
                      WHERE periodo_inicio IS NULL OR periodo_inicio = ''
                         OR periodo_fin    IS NULL OR periodo_fin    = '';",
                    transaction: transaction) > 0;
            }

            // FASE 752: Soltar triggers obsoletos ÚNICAMENTE si efectivamente se ejecutará backfill DML
            if (requiereBackfill)
            {
                LoggerHelper.Info("[V65] Paso 3: Soltando triggers CR-SQLite obsoletos previo a backfill...");
                PeriodoVentanaCrdtRepair.DropStaleCrsqlTriggers(connection, transaction);

                // Paso 4: Backfill canónico respetando regla V56 para 2026-02 y resolviendo campos nulos individuales (FASE 753)
                LoggerHelper.Info("[V65] Paso 4: Backfill parametrizado de fechas en periodos...");
                var datos = new[]
                {
                    new { Id = 1, Ini = "2026-01-12", Fin = "2026-06-19" },
                    new { Id = 2, Ini = "2026-08-11", Fin = "2026-12-12" },
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

                LoggerHelper.Info("[V65] Migración V65 de reparación completada exitosamente con backfill.");
            }
            else
            {
                LoggerHelper.Info("[V65] Columnas y fechas completas: sin drop de triggers.");
            }

            // FASE 753/757: Replicar regla V56 SIEMPRE (independiente del backfill) por si 2026-02 retenía '2026-08-15'
            connection.Execute(
                @"UPDATE periodos
                     SET periodo_inicio = '2026-08-11'
                   WHERE (idperiodo = 2 OR periodo = '2026-02')
                     AND (periodo_inicio = '2026-08-15' OR periodo_inicio IS NULL OR periodo_inicio = '');",
                transaction: transaction);
            LoggerHelper.Info("[V65] Regla V56 verificada para 2026-02.");
        }
    }
}
