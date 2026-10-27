using System;
using System.Collections.Generic;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V78: Agrega la tabla ag_filtro_seleccion e índice único plano para persistencia de filtros (Fase 1315 / 1391 / 1409).
    /// </summary>
    public sealed class MigrationV78_AddAgFiltroSeleccion : IMigration, IMigrationTolerante
    {
        public int Version => 78;
        public string Name => "AddAgFiltroSeleccion";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction = null)
        {
            LoggerHelper.Info("[V78] Verificando e incorporando tabla ag_filtro_seleccion...");

            int existe = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_filtro_seleccion';",
                transaction: transaction);

            if (existe == 0)
            {
                connection.Execute(@"CREATE TABLE IF NOT EXISTS ag_filtro_seleccion (
                    id        INTEGER PRIMARY KEY AUTOINCREMENT,
                    subunidad TEXT    NOT NULL DEFAULT '',
                    tipo      TEXT    NOT NULL DEFAULT '',
                    refid     INTEGER NOT NULL DEFAULT 0,
                    semestre  TEXT    NOT NULL DEFAULT '',
                    grupo     TEXT    NOT NULL DEFAULT '',
                    modalidad TEXT    NOT NULL DEFAULT '',
                    creado    TEXT    NOT NULL DEFAULT (datetime('now'))
                );", transaction: transaction);
            }

            // Normalización defensiva de nulos antes de indexar
            foreach (var col in new[] { "subunidad", "tipo", "semestre", "grupo", "modalidad" })
            {
                connection.Execute(
                    $"UPDATE ag_filtro_seleccion SET {col}='' WHERE {col} IS NULL;",
                    transaction: transaction);
            }

            // Deduplicación previa para garantizar creación limpia del índice único
            connection.Execute(@"DELETE FROM ag_filtro_seleccion WHERE id NOT IN
                (SELECT MAX(id) FROM ag_filtro_seleccion GROUP BY subunidad, tipo, refid);",
                transaction: transaction);

            // Índice único plano
            connection.Execute(@"CREATE UNIQUE INDEX IF NOT EXISTS ux_ag_filtro_seleccion
                ON ag_filtro_seleccion(subunidad, tipo, refid);",
                transaction: transaction);

            LoggerHelper.Info("[V78] Tabla ag_filtro_seleccion e índice ux_ag_filtro_seleccion listos.");
        }
    }
}
