using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V75: siembra bloques canónicos de domingo (F1188), coherentes con AgRanuras: DOM = MANANA.
    /// </summary>
    public sealed class MigrationV75_SembrarBloquesDomingo : IMigration, IMigrationTolerante
    {
        public int Version => 75;
        public string Name => "SembrarBloquesDomingo";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            try
            {
                int existe = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_bloque_horario';",
                    transaction: transaction);
                if (existe == 0) return;

                const string sql = @"
INSERT INTO ag_bloque_horario (dia_semana, hora_inicio, hora_fin, duracion_horas, franja, activo, subunidad)
SELECT 7, @Ini, @Fin, @Dur, @Franja, 1, ''
WHERE NOT EXISTS (
    SELECT 1 FROM ag_bloque_horario
    WHERE dia_semana = 7 AND ABS(duracion_horas - @Dur) < 0.001 AND COALESCE(subunidad, '') = ''
);";
                var semilla = new[]
                {
                    (Ini: "08:00", Fin: "12:00", Dur: 4.0, Franja: "MANANA"),
                    (Ini: "08:00", Fin: "17:00", Dur: 8.0, Franja: "MANANA")
                };

                int insertados = 0;
                foreach (var b in semilla)
                {
                    insertados += connection.Execute(sql, new { b.Ini, b.Fin, b.Dur, b.Franja }, transaction: transaction);
                }
                if (insertados > 0)
                {
                    LoggerHelper.Info($"[V75] Bloques de domingo insertados: {insertados}");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V75] Sembrado de bloques de domingo omitido: {ex.Message}");
            }
        }
    }
}
