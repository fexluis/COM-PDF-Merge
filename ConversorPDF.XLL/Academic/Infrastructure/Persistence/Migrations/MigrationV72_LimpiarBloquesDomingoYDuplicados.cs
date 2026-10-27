using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V72: Depura bloques residuales de domingo (dia_semana = 7) y variantes de 8h duplicadas en NOCHE en ag_bloque_horario (Fase 1009).
    /// </summary>
    public sealed class MigrationV72_LimpiarBloquesDomingoYDuplicados : IMigration, IMigrationTolerante
    {
        public int Version => 72;
        public string Name => "LimpiarBloquesDomingoYDuplicados";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            try
            {
                int existeTablaBloques = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_bloque_horario';",
                    transaction: transaction);

                if (existeTablaBloques > 0)
                {
                    int eliminados = connection.Execute(@"
DELETE FROM ag_bloque_horario
WHERE COALESCE(subunidad, '') = ''
  AND (dia_semana = 7 OR (duracion_horas > 4 AND franja = 'NOCHE'));", transaction: transaction);

                    LoggerHelper.Info($"[V72] Bloques de domingo y NOCHE 8h depurados: {eliminados} registros eliminados.");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V72] Error al depurar bloques de domingo y duplicados: {ex.Message}");
            }
        }
    }
}
