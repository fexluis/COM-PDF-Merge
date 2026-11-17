using System;
using System.Collections.Generic;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// V84: Agrega la columna 'vacantes_consentidas' a 'ag_corrida' para tipar el consentimiento de vacantes (Fase 1885).
    /// Migración aditiva y tolerante: no muta datos existentes ni bloquea lecturas previas.
    /// </summary>
    public sealed class MigrationV84_AddAgCorridaVacantesConsentidas : IMigration, IMigrationTolerante
    {
        public int Version => 84;
        public string Name => "AddAgCorridaVacantesConsentidas";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction = null)
        {
            AgregarColumnaSiNoExiste(connection, transaction, "ag_corrida", "vacantes_consentidas", "TEXT NULL");
        }

        private static void AgregarColumnaSiNoExiste(IDbConnection conn, IDbTransaction trans, string tabla, string columna, string definicion)
        {
            var estadoTabla = SqliteHelper.TablaExiste(conn, tabla, trans);
            switch (estadoTabla)
            {
                case ExistenciaSql.NoExiste:
                    // La tabla aún no existe en el vault: retorno normal sin error para permitir que migraciones previas la creen
                    return;

                case ExistenciaSql.Indeterminada:
                    // Fase 1892: Error de I/O o lock: lanzar para que SchemaMigrator no registre la versión y se reintente en el próximo inicio
                    throw new InvalidOperationException($"V84_ESQUEMA_INDETERMINADO: no se pudo verificar la existencia de '{tabla}'.");
            }

            var estadoColumna = SqliteHelper.ColumnaExiste(conn, tabla, columna, trans);
            if (estadoColumna == ExistenciaSql.Indeterminada)
            {
                throw new InvalidOperationException($"V84_ESQUEMA_INDETERMINADO: no se pudo verificar la existencia de la columna '{columna}' en '{tabla}'.");
            }

            if (estadoColumna == ExistenciaSql.NoExiste)
            {
                string sql = $"ALTER TABLE {tabla} ADD COLUMN {columna} {definicion};";
                conn.Execute(sql, transaction: trans);
                LoggerHelper.Info($"[V84] Columna '{columna}' agregada exitosamente a la tabla '{tabla}'.");
            }
        }
    }
}
