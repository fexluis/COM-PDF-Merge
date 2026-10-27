using System;
using System.Collections.Generic;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V54: Agrega columnas para el soporte de asignaturas vacantes en staging y producción (Fase 317).
    /// - ag_sesiones y ag_vinculaciones: es_vacante INTEGER DEFAULT 0, motivo_vacante TEXT DEFAULT '', candidatos_pool INTEGER DEFAULT 0
    /// - vinculaciones y sesiones: es_vacante INTEGER DEFAULT 0, motivo_vacante TEXT DEFAULT ''
    /// </summary>
    public class MigrationV54_AddVacanteColumns : IMigration, IMigrationTolerante
    {
        public int Version => 54;
        public string Name => "AddVacanteColumns";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V54] Verificando e incorporando columnas de vacante en tablas de staging y producción...");

            // Tablas de staging
            AgregarColumnaSiNoExiste(connection, transaction, "ag_vinculaciones", "es_vacante", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_vinculaciones", "motivo_vacante", "TEXT NOT NULL DEFAULT ''");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_vinculaciones", "candidatos_pool", "INTEGER NOT NULL DEFAULT 0");

            AgregarColumnaSiNoExiste(connection, transaction, "ag_sesiones", "es_vacante", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_sesiones", "motivo_vacante", "TEXT NOT NULL DEFAULT ''");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_sesiones", "candidatos_pool", "INTEGER NOT NULL DEFAULT 0");

            // Tablas de producción
            AgregarColumnaSiNoExiste(connection, transaction, "vinculaciones", "es_vacante", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "vinculaciones", "motivo_vacante", "TEXT NOT NULL DEFAULT ''");

            AgregarColumnaSiNoExiste(connection, transaction, "sesiones", "es_vacante", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "sesiones", "motivo_vacante", "TEXT NOT NULL DEFAULT ''");
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
                LoggerHelper.Info($"[V54] Columna '{columna}' añadida exitosamente a tabla '{tabla}'.");
            }
        }
    }
}
