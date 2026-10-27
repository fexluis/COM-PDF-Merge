using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V74: asegura la existencia de bloques horarios canónicos de 4.0 horas para días hábiles (Fase 1061).
    /// Garantiza que pregrado disponga de bloques de 4 horas para jornadas MANANA, TARDE y NOCHE sin forzar bloques de 8 horas.
    /// </summary>
    public sealed class MigrationV74_SembrarBloques4hFaltantes : IMigration, IMigrationTolerante
    {
        public int Version => 74;
        public string Name => "SembrarBloques4hFaltantes";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            try
            {
                int existe = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_bloque_horario';",
                    transaction: transaction);
                if (existe == 0) return;

                // Bloques canónicos de 4 horas por día y franja (Fase 1079: TARDE canónica 13:00-17:00 alineada con V62 y AgFranjaJornada):
                // Días 1..5: MANANA (08:00-12:00), TARDE (13:00-17:00), NOCHE (18:00-22:00)
                // Día 6: MANANA (08:00-12:00), TARDE (13:00-17:00)
                var bloquesSemilla = new[]
                {
                    // Lunes a Viernes (Días 1 a 5)
                    (1, "08:00", "12:00", "MANANA"), (1, "13:00", "17:00", "TARDE"), (1, "18:00", "22:00", "NOCHE"),
                    (2, "08:00", "12:00", "MANANA"), (2, "13:00", "17:00", "TARDE"), (2, "18:00", "22:00", "NOCHE"),
                    (3, "08:00", "12:00", "MANANA"), (3, "13:00", "17:00", "TARDE"), (3, "18:00", "22:00", "NOCHE"),
                    (4, "08:00", "12:00", "MANANA"), (4, "13:00", "17:00", "TARDE"), (4, "18:00", "22:00", "NOCHE"),
                    (5, "08:00", "12:00", "MANANA"), (5, "13:00", "17:00", "TARDE"), (5, "18:00", "22:00", "NOCHE"),
                    // Sábado (Día 6)
                    (6, "08:00", "12:00", "MANANA"), (6, "13:00", "17:00", "TARDE")
                };

                int insertados = 0;
                string sqlInsert = @"
INSERT INTO ag_bloque_horario (dia_semana, hora_inicio, hora_fin, duracion_horas, franja, activo, subunidad)
SELECT @Dia, @Ini, @Fin, 4.0, @Franja, 1, ''
WHERE NOT EXISTS (
    SELECT 1 FROM ag_bloque_horario
    WHERE dia_semana = @Dia
      AND franja = @Franja
      AND ABS(duracion_horas - 4.0) < 0.001
      AND COALESCE(subunidad, '') = ''
);";

                foreach (var (dia, ini, fin, franja) in bloquesSemilla)
                {
                    int f = connection.Execute(sqlInsert, new { Dia = dia, Ini = ini, Fin = fin, Franja = franja }, transaction: transaction);
                    insertados += f;
                }

                if (insertados > 0)
                {
                    LoggerHelper.Info($"[V74] Bloques canónicos de 4h insertados: {insertados}");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V74] Sembrado de bloques 4h omitido: {ex.Message}");
            }
        }
    }
}
