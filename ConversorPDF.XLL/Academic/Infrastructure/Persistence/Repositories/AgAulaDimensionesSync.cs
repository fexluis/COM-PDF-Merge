using System;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using CSharpFunctionalExtensions;
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
                        try
                        {
                            ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Warning($"[AgAulaDimensionesSync] MarkDirty falló tras UPDATE en '{aulaRef}': {ex.Message}");
                        }
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

        /// <summary>
        /// Sincroniza las dimensiones en ag_aula con validación explícita (Fase 1726):
        /// - Falla con Result.Failure si el aula no existe en BD o si ocurre una excepción SQL.
        /// - Retorna Result.Success(0) si el aula existe y los valores ya son coincidentes.
        /// - Retorna Result.Success(n) si se actualizaron n filas en ag_aula.
        /// </summary>
        public async Task<Result<int>> SincronizarConfirmadoAsync(string aulaRef, string jornada, int horasDia, string diasVentana, string subunidad, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(aulaRef))
            {
                return Result.Failure<int>("El aula_ref especificado es nulo o vacío.");
            }

            try
            {
                using (var conn = _connectionFactory())
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    const string sqlExiste = @"
SELECT COUNT(1)
FROM ag_aula
WHERE aula_ref = @Ref
  AND COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub;";

                    int existe = await conn.ExecuteScalarAsync<int>(new CommandDefinition(
                        sqlExiste,
                        new { Ref = aulaRef, Sub = subunidad ?? string.Empty },
                        cancellationToken: ct)).ConfigureAwait(false);

                    if (existe == 0)
                    {
                        return Result.Failure<int>($"El aula '{aulaRef}' no existe en la base de datos.");
                    }

                    const string sql = @"
UPDATE ag_aula
SET jornada      = COALESCE(NULLIF(@J, ''), jornada),
    horas_dia    = CASE WHEN @H > 0 THEN @H ELSE horas_dia END,
    dias_ventana = COALESCE(NULLIF(@D, ''), dias_ventana)
WHERE aula_ref = @Ref
  AND COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub
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
                        try
                        {
                            ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Warning($"[AgAulaDimensionesSync] MarkDirty falló tras UPDATE en '{aulaRef}': {ex.Message}");
                        }
                    }

                    return Result.Success(n);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[AgAulaDimensionesSync] Error al sincronizar aula '{aulaRef}': {ex.Message}");
                return Result.Failure<int>($"Error al sincronizar dimensiones para el aula '{aulaRef}': {ex.Message}");
            }
        }

        /// <summary>
        /// Sincroniza en base de datos de forma atómica (transacción única) un lote de aulas
        /// que presentan diferencias respecto al perfil (Fase 2306 / Fase 2308).
        /// </summary>
        public async Task<Result<int>> SincronizarLoteAsync(
            System.Collections.Generic.IEnumerable<(string AulaRef, string Jornada, int HorasDia, string DiasVentana, string Subunidad)> diffs,
            CancellationToken ct)
        {
            if (diffs == null)
            {
                return Result.Success(0);
            }

            const string sqlUpdate = @"
UPDATE ag_aula
SET jornada      = COALESCE(NULLIF(@J, ''), jornada),
    horas_dia    = CASE WHEN @H > 0 THEN @H ELSE horas_dia END,
    dias_ventana = COALESCE(NULLIF(@D, ''), dias_ventana)
WHERE aula_ref = @Ref
  AND COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub
  AND (COALESCE(jornada, '')      <> COALESCE(NULLIF(@J, ''), COALESCE(jornada, ''))
    OR COALESCE(horas_dia, -1)    <> CASE WHEN @H > 0 THEN @H ELSE COALESCE(horas_dia, -1) END
    OR COALESCE(dias_ventana, '') <> COALESCE(NULLIF(@D, ''), COALESCE(dias_ventana, '')));";

            try
            {
                using (var conn = _connectionFactory())
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (var tx = conn.BeginTransaction())
                    {
                        const string sqlExiste = @"
SELECT COUNT(1)
FROM ag_aula
WHERE aula_ref = @Ref
  AND COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub;";

                        int totalActualizadas = 0;
                        foreach (var item in diffs)
                        {
                            ct.ThrowIfCancellationRequested();

                            // Fase 2313: Verificación de existencia previa en lote
                            int existe = await conn.ExecuteScalarAsync<int>(new CommandDefinition(
                                sqlExiste,
                                new { Ref = item.AulaRef, Sub = item.Subunidad ?? string.Empty },
                                transaction: tx,
                                cancellationToken: ct)).ConfigureAwait(false);

                            if (existe == 0)
                            {
                                tx.Rollback();
                                return Result.Failure<int>($"El aula '{item.AulaRef}' no existe en la base de datos.");
                            }

                            int n = await conn.ExecuteAsync(new CommandDefinition(
                                sqlUpdate,
                                new
                                {
                                    J = item.Jornada ?? string.Empty,
                                    H = item.HorasDia,
                                    D = item.DiasVentana ?? string.Empty,
                                    Ref = item.AulaRef,
                                    Sub = item.Subunidad ?? string.Empty
                                },
                                transaction: tx,
                                cancellationToken: ct)).ConfigureAwait(false);

                            totalActualizadas += n;
                        }

                        tx.Commit();

                        // Fase 2314: MarkDirty post-commit tolerante a excepciones
                        try
                        {
                            if (totalActualizadas > 0)
                            {
                                ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();
                            }
                        }
                        catch (Exception ex)
                        {
                            LoggerHelper.Warning($"[AgAulaDimensionesSync] MarkDirty falló tras commit: {ex.Message}");
                        }

                        return Result.Success(totalActualizadas);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[AgAulaDimensionesSync] Error en SincronizarLoteAsync: {ex.Message}");
                return Result.Failure<int>($"Error al sincronizar lote de dimensiones: {ex.Message}");
            }
        }
    }
}
