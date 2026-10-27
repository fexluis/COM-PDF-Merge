using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V73: limpia dias_patron con más de 2 días en ag_aula_asignatura (Fase 1015).
    /// El motor reparte carriles de 2 días; un patrón de 3-5 días anula el reparto y sobre-divide la carga.
    /// </summary>
    public sealed class MigrationV73_LimpiarDiasPatronMultidia : IMigration, IMigrationTolerante
    {
        public int Version => 73;
        public string Name => "LimpiarDiasPatronMultidia";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            try
            {
                int existe = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_aula_asignatura';",
                    transaction: transaction);
                if (existe == 0) return;

                int afectables = connection.QueryFirstOrDefault<int>(@"
SELECT COUNT(*) FROM ag_aula_asignatura
WHERE COALESCE(TRIM(dias_patron), '') <> ''
  AND (LENGTH(dias_patron) - LENGTH(REPLACE(dias_patron, ',', '')) + 1) > 2;", transaction: transaction);

                int limpiados = connection.Execute(@"
UPDATE ag_aula_asignatura
SET dias_patron = ''
WHERE COALESCE(TRIM(dias_patron), '') <> ''
  AND (LENGTH(dias_patron) - LENGTH(REPLACE(dias_patron, ',', '')) + 1) > 2;", transaction: transaction);

                LoggerHelper.Info($"[V73] dias_patron multidía detectados={afectables}, limpiados={limpiados} (Fase 1015/1212/1222).");
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V73] Limpieza dias_patron omitida: {ex.Message}");
            }
        }
    }
}
