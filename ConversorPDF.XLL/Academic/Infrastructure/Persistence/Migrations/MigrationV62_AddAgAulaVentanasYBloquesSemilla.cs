using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V62: Agrega columnas jornada, horas_dia y dias_ventana en ag_aula y ag_arbol_nodo,
    /// y completa la semilla canónica de 42 bloques horarios (6 franjas x 7 días) de forma idempotente.
    /// </summary>
    public class MigrationV62_AddAgAulaVentanasYBloquesSemilla : IMigration, IMigrationTolerante
    {
        public int Version => 62;
        public string Name => "AddAgAulaVentanasYBloquesSemilla";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V62] Iniciando migración V62: columnas de decisión y seed de 42 bloques...");

            // 1. Columnas en ag_aula
            AgregarColumnaSiNoExiste(connection, transaction, "ag_aula", "jornada", "TEXT NOT NULL DEFAULT ''");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_aula", "horas_dia", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_aula", "dias_ventana", "TEXT NOT NULL DEFAULT ''");

            // 2. Columnas en ag_arbol_nodo
            AgregarColumnaSiNoExiste(connection, transaction, "ag_arbol_nodo", "jornada", "TEXT NOT NULL DEFAULT ''");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_arbol_nodo", "horas_dia", "INTEGER NOT NULL DEFAULT 0");
            AgregarColumnaSiNoExiste(connection, transaction, "ag_arbol_nodo", "dias_ventana", "TEXT NOT NULL DEFAULT ''");

            // 3. Semilla 42 bloques horarios (6 por día x 7 días) si existe la tabla
            int existeTablaBloques = connection.QueryFirstOrDefault<int>(@"
SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_bloque_horario';", transaction: transaction);

            if (existeTablaBloques > 0)
            {
                var plantillasBloques = new[]
                {
                    new { Ini = "08:00", Fin = "12:00", Dur = 4.0, Franja = "MANANA" },
                    new { Ini = "13:00", Fin = "17:00", Dur = 4.0, Franja = "TARDE" },
                    new { Ini = "18:00", Fin = "22:00", Dur = 4.0, Franja = "NOCHE" },
                    new { Ini = "08:00", Fin = "17:00", Dur = 8.0, Franja = "MANANA" },
                    new { Ini = "14:00", Fin = "22:00", Dur = 8.0, Franja = "TARDE" }
                };

                const string insertSql = @"
INSERT INTO ag_bloque_horario (dia_semana, hora_inicio, hora_fin, duracion_horas, franja, activo, subunidad)
SELECT @Dia, @Ini, @Fin, @Dur, @Franja, 1, ''
WHERE NOT EXISTS (
    SELECT 1 FROM ag_bloque_horario
    WHERE dia_semana = @Dia
      AND hora_inicio = @Ini
      AND hora_fin = @Fin
      AND franja = @Franja
      AND COALESCE(subunidad, '') = ''
);";

                int insertados = 0;
                for (int dia = 1; dia <= 6; dia++)   // Fase 1006: sin DOMINGO
                {
                    foreach (var b in plantillasBloques)
                    {
                        int affected = connection.Execute(insertSql, new
                        {
                            Dia = dia,
                            Ini = b.Ini,
                            Fin = b.Fin,
                            Dur = b.Dur,
                            Franja = b.Franja
                        }, transaction: transaction);
                        if (affected > 0) insertados++;
                    }
                }
                LoggerHelper.Info($"[V62] Bloques insertados: {insertados}.");
            }

            LoggerHelper.Info("[V62] Migración V62 completada.");
        }

        private static void AgregarColumnaSiNoExiste(IDbConnection conn, IDbTransaction tx, string tabla, string columna, string defTipo)
        {
            try
            {
                int existe = conn.QueryFirstOrDefault<int>($@"
SELECT COUNT(*) FROM pragma_table_info('{tabla}') WHERE name='{columna}';", transaction: tx);

                if (existe == 0)
                {
                    conn.Execute($"ALTER TABLE {tabla} ADD COLUMN {columna} {defTipo};", transaction: tx);
                    LoggerHelper.Info($"[V62] Columna '{columna}' agregada a '{tabla}'.");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V62] Advertencia agregando columna '{columna}' en '{tabla}': {ex.Message}");
            }
        }
    }
}
