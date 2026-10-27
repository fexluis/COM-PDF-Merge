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

using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>
    /// Implementación de infraestructura para consultas de vinculaciones usando Dapper.
    /// </summary>
    public class VinculacionQueryRepository : IVinculacionQueryRepository, IDisposable
    {
        private Func<IDbConnection> _connectionFactory;
        private IDisposable _rescuedContext;
        private readonly AgModoDatos _modo;

        public VinculacionQueryRepository(Func<IDbConnection> connectionFactory, AgModoDatos modo = null)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
            _modo = modo ?? AgModoDatos.Determinista;
        }

        public void Dispose()
        {
            _rescuedContext?.Dispose();
        }


        public async Task<Result<IReadOnlyCollection<SesionListItem>>> GetSesionesAsync(
            int? idSede,
            int? idTercero,
            int? idPeriodo,
            int? idPrograma,
            string modalidadPrograma,
            string subunidad,
            string semestre,
            bool soloViaticas,
            DateTime? fechaDesde,
            DateTime? fechaHasta,
            int limit,
            int offset,
            CancellationToken ct)
        {
            return await GetSesionesAsync(idSede, idTercero, idPeriodo, idPrograma, modalidadPrograma, subunidad, semestre, soloViaticas, fechaDesde, fechaHasta, limit, offset, ct, null);
        }

        public async Task<Result<IReadOnlyCollection<SesionListItem>>> GetSesionesAsync(
            int? idSede,
            int? idTercero,
            int? idPeriodo,
            int? idPrograma,
            string modalidadPrograma,
            string subunidad,
            string semestre,
            bool soloViaticas,
            DateTime? fechaDesde,
            DateTime? fechaHasta,
            int limit,
            int offset,
            CancellationToken ct,
            long? idCorrida)
        {
            var sql = $@"
                SELECT * 
                FROM {_modo.VistaSesionesDetalle} 
                WHERE 1=1";
            var parameters = new DynamicParameters();

            try
            {
                // Fase 457: Filtrar por corrida en Modo AG si se especifica
                if (idCorrida.HasValue && idCorrida.Value > 0)
                {
                    sql += " AND IdCorrida = @idCorrida";
                    parameters.Add("idCorrida", idCorrida.Value);
                }

                if (idSede.HasValue)
                {
                    sql += " AND IdSede = @idSede";
                    parameters.Add("idSede", idSede.Value);
                }

                if (idTercero.HasValue)
                {
                    sql += " AND IdTercero = @idTercero";
                    parameters.Add("idTercero", idTercero.Value);
                }

                if (idPeriodo.HasValue)
                {
                    sql += " AND IdPeriodo = @idPeriodo";
                    parameters.Add("idPeriodo", idPeriodo.Value);
                }

                if (idPrograma.HasValue)
                {
                    sql += " AND IdPrograma = @idPrograma";
                    parameters.Add("idPrograma", idPrograma.Value);
                }

                if (!string.IsNullOrEmpty(modalidadPrograma))
                {
                    var modalidadNormalizada = modalidadPrograma.Trim().ToUpperInvariant();
                    sql += " AND UPPER(TRIM(ProgramaTipo)) = @modalidadPrograma";
                    parameters.Add("modalidadPrograma", modalidadNormalizada);
                }

                if (!string.IsNullOrEmpty(subunidad))
                {
                    // Fase 13 FIX: Filtrar directamente sobre la vista para evitar fallo si falta tabla vinculaciones
                    sql += " AND Subunidad = @subunidad";
                    parameters.Add("subunidad", subunidad);
                }

                if (!string.IsNullOrEmpty(semestre))
                {
                    sql += " AND Semestre = @semestre";
                    parameters.Add("semestre", semestre);
                }

                if (soloViaticas)
                {
                    sql += " AND Viatica = 1";
                }

                if (fechaDesde.HasValue)
                {
                    sql += " AND date(FechaSesion) >= date(@fechaDesde)";
                    parameters.Add("fechaDesde", fechaDesde.Value.ToString("yyyy-MM-dd"));
                }

                if (fechaHasta.HasValue)
                {
                    sql += " AND date(FechaSesion) <= date(@fechaHasta)";
                    parameters.Add("fechaHasta", fechaHasta.Value.ToString("yyyy-MM-dd"));
                }

                sql += " ORDER BY FechaSesion DESC";
                sql += " LIMIT @limit OFFSET @offset";
                parameters.Add("limit", limit);
                parameters.Add("offset", offset);

                // Conexión compartida gestionada por el contexto del llamador (Hotfix 1.2.51.25). No disponer.
                var db = _connectionFactory();
                if (db == null)
                    return Result.Failure<IReadOnlyCollection<SesionListItem>>("La factory de conexiones retornó null.");

                if (db.State != ConnectionState.Open)
                {
                    try { db.Open(); }
                    catch (Exception ex)
                    {
                        return Result.Failure<IReadOnlyCollection<SesionListItem>>($"No se pudo abrir la conexión: {ex.Message}");
                    }
                }

                ct.ThrowIfCancellationRequested();

                // Fase 13 & 457 FIX: Verificar que la tabla base correspondiente exista (vinculaciones o ag_vinculaciones según el modo)
                string tablaVinc = _modo.TablaVinculaciones;
                int vincExists = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                    "SELECT COUNT(1) FROM sqlite_master WHERE type='table' AND name=@tablaVinc",
                    new { tablaVinc },
                    cancellationToken: ct));
                if (vincExists == 0)
                {
                    return Result.Failure<IReadOnlyCollection<SesionListItem>>($"La base de datos aún no está inicializada o la sincronización CRDT está incompleta (Falta tabla {tablaVinc}).");
                }

                // FIX FASE 10 / Fase 14 / Fase 22 / Fase 54: Verificar existencia de la columna FechaSesion y Subunidad en 'vw_SesionesDetalle' en producción
                if (!_modo.EsModoAg)
                {
                    int colExists = await db.ExecuteScalarAsync<int>(new CommandDefinition("SELECT COUNT(1) FROM sqlite_master WHERE type='view' AND name='vw_SesionesDetalle' AND sql LIKE '%FechaSesion%' AND sql LIKE '%Subunidad%'", cancellationToken: ct));
                    if (colExists == 0)
                    {
                        int termExists = await db.ExecuteScalarAsync<int>(new CommandDefinition("SELECT COUNT(1) FROM sqlite_master WHERE type='table' AND name='terminales'", cancellationToken: ct));
                        if (termExists > 0)
                        {
                            using (var tx = db.BeginTransaction())
                            {
                                try
                                {
                                    new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations.MigrationV15_AddModalidadVinculacion().Up(db, tx);
                                    new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations.MigrationV33_UpdateSesionesDetalleForTerminales().Up(db, tx);
                                    tx.Commit();
                                }
                                catch
                                {
                                    try { tx.Rollback(); } catch { }
                                    throw;
                                }
                            }
                        }
                    }
                }

                // FIX P0: Evitar NRE en Dapper por DynamicParameters vacío
                object finalParams = parameters.ParameterNames.Any() ? (object)parameters : new { };
                var items = await db.QueryAsync<SesionListItem>(new CommandDefinition(sql, finalParams, cancellationToken: ct));
                return Result.Success<IReadOnlyCollection<SesionListItem>>(items.ToList().AsReadOnly());
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<SesionListItem>>("Operación cancelada.");
            }
            catch (ObjectDisposedException)
            {
                LoggerHelper.Warning("[VinculacionQueryRepository] Conexión dispuesta, reintentando GetSesionesAsync...");
                try
                {
                    var freshCtx = SqliteHelper.GetContextForPath();
                    _rescuedContext?.Dispose();
                    _rescuedContext = freshCtx;
                    _connectionFactory = () => freshCtx.Open();

                    var newDb = freshCtx.Open();

                    // Fase 55 FIX: Re-verificar vista en conexión rescatada antes de ejecutar query
                    int colExistsRetry = await newDb.ExecuteScalarAsync<int>(new CommandDefinition("SELECT COUNT(1) FROM sqlite_master WHERE type='view' AND name='vw_SesionesDetalle' AND sql LIKE '%FechaSesion%' AND sql LIKE '%Subunidad%'", cancellationToken: ct));
                    if (colExistsRetry == 0)
                    {
                        int termExistsRetry = await newDb.ExecuteScalarAsync<int>(new CommandDefinition("SELECT COUNT(1) FROM sqlite_master WHERE type='table' AND name='terminales'", cancellationToken: ct));
                        if (termExistsRetry > 0)
                        {
                            using (var tx = newDb.BeginTransaction())
                            {
                                try
                                {
                                    new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations.MigrationV15_AddModalidadVinculacion().Up(newDb, tx);
                                    new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations.MigrationV33_UpdateSesionesDetalleForTerminales().Up(newDb, tx);
                                    tx.Commit();
                                }
                                catch { try { tx.Rollback(); } catch { } throw; }
                            }
                        }
                    }

                    object finalParams = parameters.ParameterNames.Any() ? (object)parameters : new { };
                    var items = await newDb.QueryAsync<SesionListItem>(new CommandDefinition(sql, finalParams, cancellationToken: ct));
                    return Result.Success<IReadOnlyCollection<SesionListItem>>(items.ToList().AsReadOnly());
                }
                catch (Exception exRetry)
                {
                    LoggerHelper.Error("[VinculacionQueryRepository] Error al reintentar GetSesionesAsync", exRetry);
                    return Result.Failure<IReadOnlyCollection<SesionListItem>>($"Error al consultar sesiones (reintento): {exRetry.Message}");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[VinculacionQueryRepository] Error en GetSesionesAsync", ex);
                return Result.Failure<IReadOnlyCollection<SesionListItem>>("Error al consultar sesiones: " + ex.Message);
            }
        }
    }
}
