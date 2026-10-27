using System;
using System.Collections.Generic;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V70: Agrega la columna 'vacantes' a la tabla 'ag_corrida' (Fase 956).
    /// Asegura consistencia con el modelo AgCorrida y previene errores 'no such column'.
    /// </summary>
    public sealed class MigrationV70_AddAgCorridaVacantes : IMigration, IMigrationTolerante
    {
        public int Version => 70;
        public string Name => "AddAgCorridaVacantes";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            AgregarColumnaSiNoExiste(connection, transaction, "ag_corrida", "vacantes", "INTEGER NOT NULL DEFAULT 0");
        }

        private static void AgregarColumnaSiNoExiste(IDbConnection conn, IDbTransaction trans, string tabla, string columna, string definicion)
        {
            int tablaExiste = conn.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@tabla;",
                new { tabla }, transaction: trans);

            if (tablaExiste == 0) return;

            var info = conn.Query($"PRAGMA table_info({tabla});", transaction: trans);
            bool colExiste = false;
            foreach (var row in info)
            {
                var dict = (IDictionary<string, object>)row;
                if (dict.TryGetValue("name", out var val) && string.Equals(val?.ToString(), columna, StringComparison.OrdinalIgnoreCase))
                {
                    colExiste = true;
                    break;
                }
            }

            if (!colExiste)
            {
                string sql = $"ALTER TABLE {tabla} ADD COLUMN {columna} {definicion};";
                conn.Execute(sql, transaction: trans);
                LoggerHelper.Info($"[V70] Columna '{columna}' agregada exitosamente a la tabla '{tabla}'.");
            }
        }
    }
}
