using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>
    /// Implementación Dapper del repositorio del pool docente-asignatura para SQLite con soporte CRDT.
    /// </summary>
    public class TerceroAsignaturaRepository : ITerceroAsignaturaRepository
    {
        private const string SelectBase = @"
SELECT 
    id as Id,
    idtercero as IdTercero,
    idprograma as IdPrograma,
    idasignatura as IdAsignatura,
    semestre as Semestre,
    activo as Activo,
    prioridad as Prioridad,
    veces_asignada as VecesAsignada,
    fecha_ultima_asignacion as FechaUltimaAsignacion,
    origen as Origen,
    observaciones as Observaciones,
    subunidad as Subunidad,
    created_at as CreatedAt,
    updated_at as UpdatedAt
FROM terceros_asignaturas";

        private readonly Func<IDbConnection> _connectionFactory;
        private readonly IDbTransaction _transaction;
        private readonly bool _suppressMarkDirty;

        public TerceroAsignaturaRepository(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        public TerceroAsignaturaRepository(IDbConnection connection, IDbTransaction transaction = null, bool suppressMarkDirty = false)
            : this(() => connection)
        {
            _transaction = transaction;
            _suppressMarkDirty = suppressMarkDirty;
        }

        public async Task<Result<IReadOnlyCollection<TerceroAsignatura>>> GetAllAsync(CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string sql = $"{SelectBase} ORDER BY idprograma, semestre, idasignatura, prioridad DESC";
                    var items = await db.QueryAsync<TerceroAsignatura>(new CommandDefinition(sql, cancellationToken: ct));
                    return Result.Success<IReadOnlyCollection<TerceroAsignatura>>(items.ToList().AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<TerceroAsignatura>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[TerceroAsignaturaRepository] Error en GetAllAsync", ex);
                return Result.Failure<IReadOnlyCollection<TerceroAsignatura>>($"Error al consultar pool: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<TerceroAsignatura>>> GetByTerceroAsync(int idTercero, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string sql = $"{SelectBase} WHERE idtercero = @idTercero ORDER BY idprograma, semestre, idasignatura";
                    var items = await db.QueryAsync<TerceroAsignatura>(new CommandDefinition(sql, new { idTercero }, cancellationToken: ct));
                    return Result.Success<IReadOnlyCollection<TerceroAsignatura>>(items.ToList().AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<TerceroAsignatura>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[TerceroAsignaturaRepository] Error en GetByTerceroAsync (Tercero {idTercero})", ex);
                return Result.Failure<IReadOnlyCollection<TerceroAsignatura>>($"Error al consultar pool por docente: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<TerceroAsignatura>>> GetByAsignaturaAsync(int idAsignatura, string semestre, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string sql = string.IsNullOrWhiteSpace(semestre)
                        ? $"{SelectBase} WHERE idasignatura = @idAsignatura AND activo = 1 ORDER BY prioridad DESC, veces_asignada DESC"
                        : $"{SelectBase} WHERE idasignatura = @idAsignatura AND (semestre = @semestre OR semestre = '') AND activo = 1 ORDER BY prioridad DESC, veces_asignada DESC";

                    string cleanSemestre = (semestre ?? string.Empty).Trim().ToUpperInvariant();
                    var items = await db.QueryAsync<TerceroAsignatura>(new CommandDefinition(sql, new { idAsignatura, semestre = cleanSemestre }, cancellationToken: ct));
                    return Result.Success<IReadOnlyCollection<TerceroAsignatura>>(items.ToList().AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<TerceroAsignatura>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[TerceroAsignaturaRepository] Error en GetByAsignaturaAsync (Asignatura {idAsignatura})", ex);
                return Result.Failure<IReadOnlyCollection<TerceroAsignatura>>($"Error al consultar pool por asignatura: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<TerceroAsignatura>>> GetPoolAsync(int idAsignatura, int idPrograma, string semestre, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string cleanSemestre = (semestre ?? string.Empty).Trim().ToUpperInvariant();
                    string sql = $"{SelectBase} WHERE idasignatura = @idAsignatura AND idprograma = @idPrograma " +
                                 "AND (semestre = @semestre OR semestre = '') AND activo = 1 " +
                                 "ORDER BY prioridad DESC, veces_asignada DESC";

                    var items = await db.QueryAsync<TerceroAsignatura>(new CommandDefinition(sql, new { idAsignatura, idPrograma, semestre = cleanSemestre }, cancellationToken: ct));
                    return Result.Success<IReadOnlyCollection<TerceroAsignatura>>(items.ToList().AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<TerceroAsignatura>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[TerceroAsignaturaRepository] Error en GetPoolAsync (Asignatura {idAsignatura}, Programa {idPrograma})", ex);
                return Result.Failure<IReadOnlyCollection<TerceroAsignatura>>($"Error al consultar pool: {ex.Message}");
            }
        }

        public async Task<Result<IReadOnlyCollection<TerceroAsignatura>>> GetByProgramaAsync(int idPrograma, string semestre, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string cleanSemestre = (semestre ?? string.Empty).Trim().ToUpperInvariant();
                    string sql = string.IsNullOrWhiteSpace(cleanSemestre)
                        ? $"{SelectBase} WHERE idprograma = @idPrograma ORDER BY semestre, idasignatura, prioridad DESC"
                        : $"{SelectBase} WHERE idprograma = @idPrograma AND (semestre = @semestre OR semestre = '') ORDER BY idasignatura, prioridad DESC";

                    var items = await db.QueryAsync<TerceroAsignatura>(new CommandDefinition(sql, new { idPrograma, semestre = cleanSemestre }, cancellationToken: ct));
                    return Result.Success<IReadOnlyCollection<TerceroAsignatura>>(items.ToList().AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<TerceroAsignatura>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[TerceroAsignaturaRepository] Error en GetByProgramaAsync (Programa {idPrograma})", ex);
                return Result.Failure<IReadOnlyCollection<TerceroAsignatura>>($"Error al consultar pool por programa: {ex.Message}");
            }
        }

        public async Task<Result<TerceroAsignatura>> GetByIdAsync(long id, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string sql = $"{SelectBase} WHERE id = @id LIMIT 1";
                    var item = await db.QuerySingleOrDefaultAsync<TerceroAsignatura>(new CommandDefinition(sql, new { id }, cancellationToken: ct));
                    if (item == null)
                    {
                        return Result.Failure<TerceroAsignatura>($"No se encontró la habilitación con ID {id}");
                    }

                    return Result.Success(item);
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<TerceroAsignatura>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[TerceroAsignaturaRepository] Error en GetByIdAsync ({id})", ex);
                return Result.Failure<TerceroAsignatura>($"Error al consultar habilitación por ID: {ex.Message}");
            }
        }

        public async Task<Result<TerceroAsignatura>> GetByKeyAsync(int idTercero, int idPrograma, int idAsignatura, string semestre, CancellationToken ct)
        {
            try
            {
                long id = TerceroAsignaturaKey.Calcular(idTercero, idPrograma, idAsignatura, semestre);
                return await GetByIdAsync(id, ct);
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<TerceroAsignatura>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[TerceroAsignaturaRepository] Error en GetByKeyAsync", ex);
                return Result.Failure<TerceroAsignatura>($"Error al consultar habilitación por clave: {ex.Message}");
            }
        }

        public async Task<Result> UpsertAsync(TerceroAsignatura item, CancellationToken ct)
        {
            if (item == null) return Result.Failure("El elemento no puede ser nulo.");

            try
            {
                if (item.Id == 0)
                {
                    item.Id = TerceroAsignaturaKey.Calcular(item.IdTercero, item.IdPrograma, item.IdAsignatura, item.Semestre);
                }

                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    // Verificación de colisión FNV-1a
                    string checkSql = "SELECT idtercero, idprograma, idasignatura, semestre FROM terceros_asignaturas WHERE id = @Id LIMIT 1";
                    var existente = await db.QuerySingleOrDefaultAsync(new CommandDefinition(checkSql, new { item.Id }, cancellationToken: ct));
                    if (existente != null)
                    {
                        string claveExistente = TerceroAsignaturaKey.ConstruirClave((int)existente.idtercero, (int)existente.idprograma, (int)existente.idasignatura, (string)existente.semestre);
                        string claveNueva = TerceroAsignaturaKey.ConstruirClave(item.IdTercero, item.IdPrograma, item.IdAsignatura, item.Semestre);
                        if (!string.Equals(claveExistente, claveNueva, StringComparison.OrdinalIgnoreCase))
                        {
                            return Result.Failure($"Colisión crítica de clave primaria detectada para el ID {item.Id}: {claveExistente} vs {claveNueva}");
                        }
                    }

                    // Recomputar veces_asignada con COUNT(DISTINCT) si no viene precalculado
                    if (item.VecesAsignada <= 0)
                    {
                        item.VecesAsignada = await ContarAsignacionesAsync(item.IdTercero, item.IdPrograma, item.IdAsignatura, item.Semestre, ct);
                    }

                    string upsertSql = @"
INSERT INTO terceros_asignaturas (
    id, idtercero, idprograma, idasignatura, semestre, activo, prioridad,
    veces_asignada, fecha_ultima_asignacion, origen, observaciones, subunidad,
    created_at, updated_at
) VALUES (
    @Id, @IdTercero, @IdPrograma, @IdAsignatura, @Semestre, @Activo, @Prioridad,
    @VecesAsignada, @FechaUltimaAsignacion, @Origen, @Observaciones, @Subunidad,
    COALESCE(NULLIF(@CreatedAt, ''), datetime('now')), datetime('now')
)
ON CONFLICT(id) DO UPDATE SET
    activo = excluded.activo,
    prioridad = excluded.prioridad,
    veces_asignada = MAX(excluded.veces_asignada, terceros_asignaturas.veces_asignada),
    fecha_ultima_asignacion = CASE
        WHEN excluded.fecha_ultima_asignacion = '' THEN terceros_asignaturas.fecha_ultima_asignacion
        ELSE excluded.fecha_ultima_asignacion END,
    observaciones = excluded.observaciones,
    subunidad = excluded.subunidad,
    updated_at = datetime('now');";

                    await db.ExecuteAsync(new CommandDefinition(upsertSql, item, cancellationToken: ct));
                    return Result.Success();
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[TerceroAsignaturaRepository] Error en UpsertAsync (ID {item.Id})", ex);
                return Result.Failure($"Error al guardar en el pool: {ex.Message}");
            }
        }

        public Result Upsert(TerceroAsignatura item, IDbTransaction transaction = null)
        {
            if (item == null) return Result.Failure("El elemento no puede ser nulo.");

            try
            {
                if (item.Id == 0)
                {
                    item.Id = TerceroAsignaturaKey.Calcular(item.IdTercero, item.IdPrograma, item.IdAsignatura, item.Semestre);
                }

                var tx = transaction ?? _transaction;
                if (tx != null)
                {
                    var db = tx.Connection ?? _connectionFactory();
                    if (db.State != ConnectionState.Open) db.Open();
                    return ExecuteUpsertInternal(db, tx, item);
                }
                else
                {
                    using (var db = _connectionFactory())
                    {
                        if (db.State != ConnectionState.Open) db.Open();
                        return ExecuteUpsertInternal(db, null, item);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[TerceroAsignaturaRepository] Error en Upsert (ID {item.Id})", ex);
                return Result.Failure($"Error al guardar en el pool: {ex.Message}");
            }
        }

        private Result ExecuteUpsertInternal(IDbConnection db, IDbTransaction tx, TerceroAsignatura item)
        {
            // Verificación de colisión FNV-1a
            string checkSql = "SELECT idtercero, idprograma, idasignatura, semestre FROM terceros_asignaturas WHERE id = @Id LIMIT 1";
            var existente = db.QuerySingleOrDefault(checkSql, new { item.Id }, transaction: tx);
            if (existente != null)
            {
                string claveExistente = TerceroAsignaturaKey.ConstruirClave((int)existente.idtercero, (int)existente.idprograma, (int)existente.idasignatura, (string)existente.semestre);
                string claveNueva = TerceroAsignaturaKey.ConstruirClave(item.IdTercero, item.IdPrograma, item.IdAsignatura, item.Semestre);
                if (!string.Equals(claveExistente, claveNueva, StringComparison.OrdinalIgnoreCase))
                {
                    return Result.Failure($"Colisión crítica de clave primaria detectada para el ID {item.Id}: {claveExistente} vs {claveNueva}");
                }
            }

            // Recomputar veces_asignada con COUNT(DISTINCT) si no viene precalculado
            if (item.VecesAsignada <= 0)
            {
                item.VecesAsignada = ContarAsignaciones(item.IdTercero, item.IdPrograma, item.IdAsignatura, item.Semestre);
            }

            string upsertSql = @"
INSERT INTO terceros_asignaturas (
    id, idtercero, idprograma, idasignatura, semestre, activo, prioridad,
    veces_asignada, fecha_ultima_asignacion, origen, observaciones, subunidad,
    created_at, updated_at
) VALUES (
    @Id, @IdTercero, @IdPrograma, @IdAsignatura, @Semestre, @Activo, @Prioridad,
    @VecesAsignada, @FechaUltimaAsignacion, @Origen, @Observaciones, @Subunidad,
    COALESCE(NULLIF(@CreatedAt, ''), datetime('now')), datetime('now')
)
ON CONFLICT(id) DO UPDATE SET
    activo = excluded.activo,
    prioridad = excluded.prioridad,
    veces_asignada = MAX(excluded.veces_asignada, terceros_asignaturas.veces_asignada),
    fecha_ultima_asignacion = CASE
        WHEN excluded.fecha_ultima_asignacion = '' THEN terceros_asignaturas.fecha_ultima_asignacion
        ELSE excluded.fecha_ultima_asignacion END,
    observaciones = excluded.observaciones,
    subunidad = excluded.subunidad,
    updated_at = datetime('now');";

            db.Execute(upsertSql, item, transaction: tx);
            return Result.Success();
        }

        public async Task<int> ContarAsignacionesAsync(int idTercero, int idPrograma, int idAsignatura, string semestre, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string cleanSemestre = (semestre ?? string.Empty).Trim().ToUpperInvariant();
                    string countSql = @"
SELECT COUNT(DISTINCT id) FROM vinculaciones 
WHERE idtercero = @idTercero 
  AND idprograma = @idPrograma 
  AND idasignatura = @idAsignatura 
  AND (semestre = @cleanSemestre OR @cleanSemestre = '' OR semestre IS NULL)";

                    return await db.ExecuteScalarAsync<int>(new CommandDefinition(countSql, new { idTercero, idPrograma, idAsignatura, cleanSemestre }, cancellationToken: ct));
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[TerceroAsignaturaRepository] Error contando asignaciones: {ex.Message}");
                return 0;
            }
        }

        public int ContarAsignaciones(int idTercero, int idPrograma, int idAsignatura, string semestre)
        {
            try
            {
                var tx = _transaction;
                if (tx != null)
                {
                    var db = tx.Connection ?? _connectionFactory();
                    if (db.State != ConnectionState.Open) db.Open();
                    return ExecuteContarAsignacionesInternal(db, tx, idTercero, idPrograma, idAsignatura, semestre);
                }
                else
                {
                    using (var db = _connectionFactory())
                    {
                        if (db.State != ConnectionState.Open) db.Open();
                        return ExecuteContarAsignacionesInternal(db, null, idTercero, idPrograma, idAsignatura, semestre);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[TerceroAsignaturaRepository] Error contando asignaciones síncrono: {ex.Message}");
                return 0;
            }
        }

        private int ExecuteContarAsignacionesInternal(IDbConnection db, IDbTransaction tx, int idTercero, int idPrograma, int idAsignatura, string semestre)
        {
            string cleanSemestre = (semestre ?? string.Empty).Trim().ToUpperInvariant();
            string countSql = @"
SELECT COUNT(DISTINCT id) FROM vinculaciones 
WHERE idtercero = @idTercero 
  AND idprograma = @idPrograma 
  AND idasignatura = @idAsignatura 
  AND (semestre = @cleanSemestre OR @cleanSemestre = '' OR semestre IS NULL)";

            return db.ExecuteScalar<int>(countSql, new { idTercero, idPrograma, idAsignatura, cleanSemestre }, transaction: tx);
        }

        public async Task<Result> DeleteAsync(long id, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string sql = "DELETE FROM terceros_asignaturas WHERE id = @id";
                    int rows = await db.ExecuteAsync(new CommandDefinition(sql, new { id }, cancellationToken: ct));
                    if (rows == 0)
                    {
                        return Result.Failure($"No se encontró ningún registro para eliminar con ID {id}");
                    }

                    return Result.Success();
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[TerceroAsignaturaRepository] Error en DeleteAsync (ID {id})", ex);
                return Result.Failure($"Error al eliminar del pool: {ex.Message}");
            }
        }
    }
}
