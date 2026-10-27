using System;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>
    /// Sincroniza en tiempo real las dimensiones de decisión (Jornada, Horas/Día, Días de Ventana)
    /// seleccionadas en la interfaz de usuario directamente con los registros de ag_aula en base de datos (Fase 1193 / Fase 3).
    /// </summary>
    public sealed class AgAulaDimensionesSync
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public AgAulaDimensionesSync(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory;
        }

        /// <summary>
        /// Actualiza las columnas jornada, horas_dia y dias_ventana en ag_aula para el aula_ref y subunidad indicados,
        /// únicamente si los valores en base de datos difieren de los seleccionados (Fase 1198).
        /// </summary>
        public async Task<int> SincronizarAsync(string aulaRef, string jornada, int horasDia, string diasVentana, string subunidad, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(aulaRef))
            {
                return 0;
            }

            try
            {
                using (var conn = _connectionFactory())
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    const string sql = @"
UPDATE ag_aula
SET jornada     = COALESCE(NULLIF(@J, ''), jornada),      -- Fase 1246: no borrar dimensión legacy
    horas_dia   = CASE WHEN @H > 0 THEN @H ELSE horas_dia END,
    dias_ventana = COALESCE(NULLIF(@D, ''), dias_ventana)
WHERE aula_ref = @Ref
  -- Fase 1238: aísla por subunidad; tolera filas legacy vacías sin barrer otras subunidades
  AND COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub
  -- Fase 1248: el predicado compara contra el valor EFECTIVO, no contra el crudo
  AND (COALESCE(jornada, '')      <> COALESCE(NULLIF(@J, ''), COALESCE(jornada, ''))
    OR COALESCE(horas_dia, -1)    <> CASE WHEN @H > 0 THEN @H ELSE COALESCE(horas_dia, -1) END
    OR COALESCE(dias_ventana, '') <> COALESCE(NULLIF(@D, ''), COALESCE(dias_ventana, '')));";

                    int n = await conn.ExecuteAsync(new CommandDefinition(
                        sql,
                        new
                        {
                            J = jornada ?? string.Empty,
                            H = horasDia,
                            D = diasVentana ?? string.Empty,
                            Ref = aulaRef,
                            Sub = subunidad ?? string.Empty
                        },
                        cancellationToken: ct)).ConfigureAwait(false);

                    if (n > 0)
                    {
                        ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();
                    }

                    return n;
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[AgAulaDimensionesSync] Error al sincronizar aula '{aulaRef}': {ex.Message}");
                return 0;
            }
        }
    }
}
