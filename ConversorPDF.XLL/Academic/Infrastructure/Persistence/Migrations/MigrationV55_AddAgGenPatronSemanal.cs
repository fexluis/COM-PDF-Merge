using System;
using System.Collections.Generic;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V55: Agrega columnas para soporte de patrones multi-día y siembra franja Lunes (Fase 382 / Fase 395 / Fase 402).
    /// - ag_gen: indice_dia_patron, patron_semanal, dias_patron, clave_grupo, horas_gen
    /// - ag_aula_asignatura: dias_patron
    /// - ag_bloque_horario: seed de Lunes 18:00-22:00 NOCHE si no existe
    /// </summary>
    public class MigrationV55_AddAgGenPatronSemanal : IMigration, IMigrationTolerante
    {
        public int Version => 55;
        public string Name => "AddAgGenPatronSemanal";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V55] Verificando e incorporando columnas de patrón multi-día y franja Lunes...");

            // Columnas en ag_gen
            AgregarColumnaSiNoExiste(connection, transaction, "ag_gen", "indice_dia_patron", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_gen", "patron_semanal", "TEXT NOT NULL DEFAULT ''");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_gen", "dias_patron", "TEXT NOT NULL DEFAULT ''");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_gen", "clave_grupo", "TEXT NOT NULL DEFAULT ''");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_gen", "horas_gen", "REAL NOT NULL DEFAULT 4.0");

            // Columnas en ag_aula_asignatura
            AgregarColumnaSiNoExiste(connection, transaction, "ag_aula_asignatura", "dias_patron", "TEXT NOT NULL DEFAULT ''");

            // Seed franja Lunes en ag_bloque_horario si no existe (Fase 395)
            SembrarFranjaLunesSiNoExiste(connection, transaction);
        }

        private static void SembrarFranjaLunesSiNoExiste(IDbConnection conn, IDbTransaction trans)
        {
            int tablaExiste = conn.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_bloque_horario';",
                transaction: trans);

            if (tablaExiste == 0) return;

            int lunesExiste = conn.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM ag_bloque_horario WHERE dia_semana = 1 AND activo = 1;",
                transaction: trans);

            if (lunesExiste == 0)
            {
                conn.Execute(@"
                    INSERT INTO ag_bloque_horario (dia_semana, hora_inicio, hora_fin, duracion_horas, franja, activo, subunidad)
                    VALUES (1, '18:00', '22:00', 4.0, 'NOCHE', 1, '');",
                    transaction: trans);
                LoggerHelper.Info("[V55] Franja LUNES 18:00-22:00 sembrada exitosamente en ag_bloque_horario.");
            }
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
                LoggerHelper.Info($"[V55] Columna '{columna}' añadida exitosamente a tabla '{tabla}'.");
            }
        }
    }
}
