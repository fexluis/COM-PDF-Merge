using System;
using System.Collections.Generic;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V77: Agrega la columna isPernoctaUltimoDia en tablas de vinculaciones (vinculaciones y ag_vinculaciones) (Fase 1279).
    /// Corrige la asimetría donde ag_vinculaciones carecía del campo de pernocta, ocasionando fallos SQL
    /// en el repositorio académico al operar en modo AG (staging).
    /// </summary>
    public sealed class MigrationV77_AddPernoctaAgVinculaciones : IMigration, IMigrationTolerante
    {
        public int Version => 77;
        public string Name => "AddPernoctaAgVinculaciones";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V77] Verificando e incorporando columna isPernoctaUltimoDia en vinculaciones y ag_vinculaciones...");

            AgregarColumnaSiNoExiste(connection, transaction, "vinculaciones", "isPernoctaUltimoDia", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_vinculaciones", "isPernoctaUltimoDia", "INTEGER NOT NULL DEFAULT 0");
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
                LoggerHelper.Info($"[V77] Columna '{columna}' añadida exitosamente a tabla '{tabla}'.");
            }
        }
    }
}
