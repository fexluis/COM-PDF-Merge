using System;
using System.Collections.Generic;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V57: Repara la tabla salvedades_conflictos y ag_salvedades_conflictos agregando las columnas faltantes idsesion_a e idsesion_b (Fase 459).
    /// </summary>
    public class MigrationV57_RepairSalvedadesConflictos : IMigration, IMigrationTolerante
    {
        public int Version => 57;
        public string Name => "RepairSalvedadesConflictos";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V57] Reparando columnas idsesion_a e idsesion_b en salvedades_conflictos...");

            AgregarColumnaSiNoExiste(connection, transaction, "salvedades_conflictos", "idsesion_a", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "salvedades_conflictos", "idsesion_b", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_salvedades_conflictos", "idsesion_a", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_salvedades_conflictos", "idsesion_b", "INTEGER NOT NULL DEFAULT 0");

            LoggerHelper.Info("[V57] Columnas de salvedades_conflictos verificadas y reparadas exitosamente.");
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
                conn.Execute($"ALTER TABLE {tabla} ADD COLUMN {columna} {definicion};", transaction: trans);
                LoggerHelper.Info($"[V57] Columna '{columna}' agregada a tabla '{tabla}'.");
            }
        }
    }
}
