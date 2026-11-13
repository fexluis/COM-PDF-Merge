using System;
using System.Collections.Generic;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V80: Incorpora la columna 'origen_patron' en ag_aula_asignatura (Fase 1735).
    /// Permite distinguir patrones AUTOMATICO, FIJO y LEGADO sin alterar datos existentes.
    /// </summary>
    public sealed class MigrationV80_AddOrigenPatronAgAulaAsignatura : IMigration, IMigrationTolerante
    {
        public int Version => 80;
        public string Name => "AddOrigenPatronAgAulaAsignatura";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction = null)
        {
            LoggerHelper.Info("[V80] Verificando e incorporando columna 'origen_patron' en ag_aula_asignatura...");

            int tablaExiste = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_aula_asignatura';",
                transaction: transaction);

            if (tablaExiste == 0) return;

            var info = connection.Query("PRAGMA table_info(ag_aula_asignatura);", transaction: transaction);
            bool colExiste = false;
            foreach (var row in info)
            {
                var dict = (IDictionary<string, object>)row;
                if (dict.TryGetValue("name", out var val) && string.Equals(val?.ToString(), "origen_patron", StringComparison.OrdinalIgnoreCase))
                {
                    colExiste = true;
                    break;
                }
            }

            if (!colExiste)
            {
                connection.Execute(
                    "ALTER TABLE ag_aula_asignatura ADD COLUMN origen_patron TEXT NOT NULL DEFAULT 'LEGADO';",
                    transaction: transaction);
                LoggerHelper.Info("[V80] Columna 'origen_patron' agregada exitosamente.");
            }

            // Normalizar valores nulos o vacíos a LEGADO
            connection.Execute(
                "UPDATE ag_aula_asignatura SET origen_patron = 'LEGADO' WHERE origen_patron IS NULL OR TRIM(origen_patron) = '';",
                transaction: transaction);

            LoggerHelper.Info("[V80] Migración V80 completada.");
        }
    }
}

