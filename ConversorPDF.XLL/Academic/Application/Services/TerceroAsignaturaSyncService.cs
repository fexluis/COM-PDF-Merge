using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Implementación del servicio de sincronización histórica del pool docente-asignatura.
    /// Recomputa de forma idempotente las veces asignada y la fecha de última asignación a partir de 'vinculaciones'.
    /// </summary>
    public class TerceroAsignaturaSyncService : ITerceroAsignaturaSyncService
    {
        private readonly Func<IDbConnection> _connectionFactory;
        private readonly ITerceroAsignaturaRepository _repository;

        public TerceroAsignaturaSyncService(
            Func<IDbConnection> connectionFactory,
            ITerceroAsignaturaRepository repository)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        public async Task<Result<int>> SincronizarTodoAsync(string subunidad, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string querySql = @"
SELECT 
    v.idtercero as IdTercero,
    v.idprograma as IdPrograma,
    v.idasignatura as IdAsignatura,
    COALESCE(v.semestre, '') as Semestre,
    COUNT(DISTINCT v.id) as VecesAsignada,
    MAX(v.fecha_inicio) as FechaUltimaAsignacionRaw
FROM vinculaciones v
WHERE v.idtercero > 0 AND v.idasignatura > 0
GROUP BY v.idtercero, v.idprograma, v.idasignatura, COALESCE(v.semestre, '')";

                    var historicos = (await db.QueryAsync(new CommandDefinition(querySql, cancellationToken: ct))).ToList();
                    int procesados = 0;

                    foreach (var h in historicos)
                    {
                        ct.ThrowIfCancellationRequested();

                        int idTercero = (int)h.IdTercero;
                        int idPrograma = (int)h.IdPrograma;
                        int idAsignatura = (int)h.IdAsignatura;
                        string semestre = ((string)h.Semestre ?? string.Empty).Trim().ToUpperInvariant();
                        int veces = (int)h.VecesAsignada;
                        string rawFecha = (string)h.FechaUltimaAsignacionRaw;

                        string fechaIso = NormalizarFechaIso(rawFecha);
                        long id = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, semestre);

                        var item = new TerceroAsignatura
                        {
                            Id = id,
                            IdTercero = idTercero,
                            IdPrograma = idPrograma,
                            IdAsignatura = idAsignatura,
                            Semestre = semestre,
                            Activo = 1,
                            Prioridad = 0,
                            VecesAsignada = veces,
                            FechaUltimaAsignacion = fechaIso,
                            Origen = "HISTORICO",
                            Observaciones = string.Empty,
                            Subunidad = subunidad ?? string.Empty
                        };

                        var result = await _repository.UpsertAsync(item, ct);
                        if (result.IsFailure)
                        {
                            LoggerHelper.Warning($"[TerceroAsignaturaSyncService] Fallo al sincronizar tupla ({idTercero},{idPrograma},{idAsignatura},{semestre}): {result.Error}");
                            continue;
                        }

                        procesados++;
                    }

                    return Result.Success(procesados);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<int>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[TerceroAsignaturaSyncService] Error al sincronizar todo el pool", ex);
                return Result.Failure<int>($"Error en sincronización general: {ex.Message}");
            }
        }

        public async Task<Result<int>> SincronizarPorSemestreAsync(int idPrograma, string semestre, string subunidad, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string cleanSemestre = (semestre ?? string.Empty).Trim().ToUpperInvariant();
                    string querySql = @"
SELECT 
    v.idtercero as IdTercero,
    v.idprograma as IdPrograma,
    v.idasignatura as IdAsignatura,
    COALESCE(v.semestre, '') as Semestre,
    COUNT(DISTINCT v.id) as VecesAsignada,
    MAX(v.fecha_inicio) as FechaUltimaAsignacionRaw
FROM vinculaciones v
WHERE v.idprograma = @idPrograma 
  AND COALESCE(v.semestre, '') = @cleanSemestre
  AND v.idtercero > 0 
  AND v.idasignatura > 0
GROUP BY v.idtercero, v.idprograma, v.idasignatura, COALESCE(v.semestre, '')";

                    var historicos = (await db.QueryAsync(new CommandDefinition(querySql, new { idPrograma, cleanSemestre }, cancellationToken: ct))).ToList();
                    int procesados = 0;

                    foreach (var h in historicos)
                    {
                        ct.ThrowIfCancellationRequested();

                        int idTercero = (int)h.IdTercero;
                        int idAsignatura = (int)h.IdAsignatura;
                        int veces = (int)h.VecesAsignada;
                        string rawFecha = (string)h.FechaUltimaAsignacionRaw;

                        string fechaIso = NormalizarFechaIso(rawFecha);
                        long id = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, cleanSemestre);

                        var item = new TerceroAsignatura
                        {
                            Id = id,
                            IdTercero = idTercero,
                            IdPrograma = idPrograma,
                            IdAsignatura = idAsignatura,
                            Semestre = cleanSemestre,
                            Activo = 1,
                            Prioridad = 0,
                            VecesAsignada = veces,
                            FechaUltimaAsignacion = fechaIso,
                            Origen = "HISTORICO",
                            Observaciones = string.Empty,
                            Subunidad = subunidad ?? string.Empty
                        };

                        var result = await _repository.UpsertAsync(item, ct);
                        if (result.IsSuccess)
                        {
                            procesados++;
                        }
                    }

                    return Result.Success(procesados);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<int>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[TerceroAsignaturaSyncService] Error al sincronizar por semestre (P:{idPrograma}, S:{semestre})", ex);
                return Result.Failure<int>($"Error en sincronización por semestre: {ex.Message}");
            }
        }

        public async Task<Result> SincronizarDocenteAsignaturaAsync(int idTercero, int idPrograma, int idAsignatura, string semestre, string subunidad, CancellationToken ct)
        {
            if (idTercero <= 0 || idAsignatura <= 0) return Result.Success();

            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string cleanSemestre = (semestre ?? string.Empty).Trim().ToUpperInvariant();
                    string querySql = @"
SELECT 
    COUNT(DISTINCT v.id) as VecesAsignada,
    MAX(v.fecha_inicio) as FechaUltimaAsignacionRaw
FROM vinculaciones v
WHERE v.idtercero = @idTercero
  AND v.idprograma = @idPrograma
  AND v.idasignatura = @idAsignatura
  AND COALESCE(v.semestre, '') = @cleanSemestre";

                    var stat = await db.QuerySingleOrDefaultAsync(new CommandDefinition(querySql, new { idTercero, idPrograma, idAsignatura, cleanSemestre }, cancellationToken: ct));
                    int veces = stat != null ? (int)stat.VecesAsignada : 0;
                    string rawFecha = stat != null ? (string)stat.FechaUltimaAsignacionRaw : string.Empty;
                    string fechaIso = NormalizarFechaIso(rawFecha);

                    long id = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, cleanSemestre);

                    var existenteResult = await _repository.GetByIdAsync(id, ct);
                    if (existenteResult.IsSuccess)
                    {
                        var existente = existenteResult.Value;
                        existente.VecesAsignada = veces;
                        existente.FechaUltimaAsignacion = fechaIso;
                        return await _repository.UpsertAsync(existente, ct);
                    }

                    var nuevo = new TerceroAsignatura
                    {
                        Id = id,
                        IdTercero = idTercero,
                        IdPrograma = idPrograma,
                        IdAsignatura = idAsignatura,
                        Semestre = cleanSemestre,
                        Activo = 1,
                        Prioridad = 0,
                        VecesAsignada = veces,
                        FechaUltimaAsignacion = fechaIso,
                        Origen = "HISTORICO",
                        Observaciones = string.Empty,
                        Subunidad = subunidad ?? string.Empty
                    };

                    return await _repository.UpsertAsync(nuevo, ct);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[TerceroAsignaturaSyncService] Error en SincronizarDocenteAsignaturaAsync (T:{idTercero}, A:{idAsignatura})", ex);
                return Result.Failure($"Error al sincronizar habilitación docente: {ex.Message}");
            }
        }

        private static string NormalizarFechaIso(string rawFecha)
        {
            if (string.IsNullOrWhiteSpace(rawFecha)) return string.Empty;

            string[] formats = { "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "yyyy/MM/dd" };
            if (DateTime.TryParseExact(rawFecha.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            {
                return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            if (DateTime.TryParse(rawFecha.Trim(), out var dtLoose))
            {
                return dtLoose.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            return rawFecha.Trim();
        }
    }
}
