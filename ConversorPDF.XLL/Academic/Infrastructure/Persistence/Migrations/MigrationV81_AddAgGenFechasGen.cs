using System;
using System.Collections.Generic;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V81: Incorpora la columna 'fechas_gen' en la tabla 'ag_gen'.
    /// Permite persistir y rehidratar el plan fechado intercalado autorizado por gen,
    /// garantizando que el expansor de sesiones y la compuerta de promoción no recurran
    /// al generador semanal legacy divergente.
    /// </summary>
    public sealed class MigrationV81_AddAgGenFechasGen : IMigration, IMigrationTolerante
    {
        public int Version => 81;
        public string Name => "AddAgGenFechasGen";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction = null)
        {
            LoggerHelper.Info("[V81] Verificando e incorporando columna 'fechas_gen' en ag_gen...");

            int tablaExiste = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_gen';",
                transaction: transaction);

            if (tablaExiste == 0) return;

            var info = connection.Query("PRAGMA table_info(ag_gen);", transaction: transaction);
            bool colExiste = false;
            foreach (var row in info)
            {
                var dict = (IDictionary<string, object>)row;
                if (dict.TryGetValue("name", out var val) && string.Equals(val?.ToString(), "fechas_gen", StringComparison.OrdinalIgnoreCase))
                {
                    colExiste = true;
                    break;
                }
            }

            if (!colExiste)
            {
                connection.Execute(
                    "ALTER TABLE ag_gen ADD COLUMN fechas_gen TEXT NULL;",
                    transaction: transaction);
                LoggerHelper.Info("[V81] Columna 'fechas_gen' agregada exitosamente en ag_gen.");
            }

            LoggerHelper.Info("[V81] Migración V81 completada.");
        }
    }
}
