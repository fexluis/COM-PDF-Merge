using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
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
            var parameters = new DynamicParameters();
            var sql = $"SELECT * FROM {_modo.VistaSesionesDetalle} WHERE 1=1";

            try
            {
                // Fase 1490: predicado único compartido con los contadores del resumen de staging
                sql += BuildWhere(parameters, idCorrida, idSede, idTercero, idPeriodo, idPrograma,
                                  modalidadPrograma, subunidad, semestre, soloViaticas, fechaDesde, fechaHasta);

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

        /// <summary>
        /// Obtiene las corridas del staging con su volumen real de vinculaciones y sesiones (Fase 1489).
        /// Incluye tanto las corridas con staging materializado en <c>ag_vinculaciones</c> como las
        /// registradas en <c>ag_corrida</c> que aún no tienen filas (se listan como "sin sesiones"),
        /// y también el staging huérfano que no tiene cabecera en <c>ag_corrida</c>.
        /// </summary>
        public async Task<Result<IReadOnlyCollection<CorridaStagingItem>>> GetCorridasStagingAsync(
            int? idPeriodo,
            int? idPrograma,
            string subunidad,
            CancellationToken ct)
        {
            try
            {
                var db = _connectionFactory();
                if (db == null)
                    return Result.Failure<IReadOnlyCollection<CorridaStagingItem>>("La factory de conexiones retornó null.");

                if (db.State != ConnectionState.Open) db.Open();
                ct.ThrowIfCancellationRequested();

                if (!await ExisteTablaAsync(db, "ag_vinculaciones", ct))
                    return Result.Failure<IReadOnlyCollection<CorridaStagingItem>>("La base de datos no tiene staging AG (falta tabla ag_vinculaciones).");

                bool tieneCorrida = await ExisteTablaAsync(db, "ag_corrida", ct);

                // Segunda rama del UNION: corridas registradas sin vinculaciones materializadas.
                // Si la tabla no existe en vaults legados, se sustituye por un SELECT vacío.
                string ramaSinStaging = tieneCorrida
                    ? @"
UNION ALL
SELECT c2.id AS IdCorrida,
       c2.idperiodo AS IdPeriodo,
       COALESCE(pe2.periodo, '') AS PeriodoNombre,
       c2.idprograma AS IdPrograma,
       COALESCE(COALESCE(pr2.abreviatura, pr2.nombreprograma), '') AS ProgramaNombre,
       COALESCE(c2.subunidad, '') AS Subunidad,
       COALESCE(c2.fecha_inicio_corrida, '') AS FechaInicioCorrida,
       COALESCE(c2.estado, 0) AS EstadoCorrida,
       0 AS TotalVinculaciones,
       0 AS TotalSesiones
FROM ag_corrida c2
LEFT JOIN periodos  pe2 ON pe2.idperiodo = c2.idperiodo
LEFT JOIN programas pr2 ON pr2.id = c2.idprograma
WHERE NOT EXISTS (SELECT 1 FROM ag_vinculaciones v2 WHERE v2.id_corrida = c2.id)
  AND (@idPeriodo  IS NULL OR c2.idperiodo  = @idPeriodo)
  AND (@idPrograma IS NULL OR c2.idprograma = @idPrograma)
  AND (@subunidad  IS NULL OR UPPER(TRIM(c2.subunidad)) = UPPER(TRIM(@subunidad)))"
                    : string.Empty;

                string sql = @"
SELECT v.id_corrida AS IdCorrida,
       MAX(v.idperiodo) AS IdPeriodo,
       COALESCE(MAX(pe.periodo), '') AS PeriodoNombre,
       MAX(v.idprograma) AS IdPrograma,
       COALESCE(MAX(COALESCE(pr.abreviatura, pr.nombreprograma)), '') AS ProgramaNombre,
       COALESCE(MAX(NULLIF(TRIM(v.subunidad), '')), '') AS Subunidad,
       COALESCE(MAX(c.fecha_inicio_corrida), '') AS FechaInicioCorrida,
       COALESCE(MAX(c.estado), 0) AS EstadoCorrida,
       COUNT(DISTINCT v.id) AS TotalVinculaciones,
       COUNT(s.id) AS TotalSesiones
FROM ag_vinculaciones v
LEFT JOIN ag_sesiones  s  ON s.idvinculacion = v.id
LEFT JOIN ag_corrida   c  ON c.id = v.id_corrida
LEFT JOIN periodos     pe ON pe.idperiodo = v.idperiodo
LEFT JOIN programas    pr ON pr.id = v.idprograma
WHERE IFNULL(v.estado, 0) <> 3
  AND (@idPeriodo  IS NULL OR v.idperiodo  = @idPeriodo)
  AND (@idPrograma IS NULL OR v.idprograma = @idPrograma)
  AND (@subunidad  IS NULL OR UPPER(TRIM(v.subunidad)) = UPPER(TRIM(@subunidad)))
GROUP BY v.id_corrida" + ramaSinStaging + @"
ORDER BY 1 DESC;";

                var items = await db.QueryAsync<CorridaStagingItem>(
                    new CommandDefinition(sql, new { idPeriodo, idPrograma, subunidad }, cancellationToken: ct));

                return Result.Success<IReadOnlyCollection<CorridaStagingItem>>(items.ToList().AsReadOnly());
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<CorridaStagingItem>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[VinculacionQueryRepository] Error en GetCorridasStagingAsync", ex);
                return Result.Failure<IReadOnlyCollection<CorridaStagingItem>>("Error al consultar corridas de staging: " + ex.Message);
            }
        }


        /// <summary>
        /// Resumen de diagnóstico del staging de una corrida (Fases 1489/1499).
        /// La consulta A describe la corrida SIN filtros de UI; la consulta B aplica el mismo predicado
        /// que el listado (BuildWhere), de modo que "lo que se cuenta" coincide con "lo que se lista".
        /// </summary>
        public async Task<Result<AgStagingResumen>> GetResumenStagingAsync(
            long idCorrida,
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
            CancellationToken ct)
        {
            if (idCorrida <= 0)
                return Result.Failure<AgStagingResumen>("Corrida inválida (id <= 0).");

            try
            {
                var db = _connectionFactory();
                if (db == null)
                    return Result.Failure<AgStagingResumen>("La factory de conexiones retornó null.");

                if (db.State != ConnectionState.Open) db.Open();
                ct.ThrowIfCancellationRequested();

                string vista = _modo.VistaSesionesDetalle;
                if (!await ExisteVistaAsync(db, vista, ct))
                    return Result.Failure<AgStagingResumen>($"No existe la vista {vista}. Recrear vistas AG (EnsureAgViews).");

                string sqlA = $@"
SELECT COUNT(*) AS SesionesCorrida,
       COUNT(DISTINCT IdVinculacion) AS VincCorrida,
       (SELECT GROUP_CONCAT(p.periodo, ', ')
          FROM (SELECT DISTINCT IdPeriodo FROM {vista} WHERE IdCorrida = @idCorrida ORDER BY IdPeriodo) d
          LEFT JOIN periodos p ON p.idperiodo = d.IdPeriodo) AS PeriodosCorrida,
       (SELECT GROUP_CONCAT(d.Subunidad, ', ')
          FROM (SELECT DISTINCT Subunidad FROM {vista}
                 WHERE IdCorrida = @idCorrida AND Subunidad IS NOT NULL AND TRIM(Subunidad) <> ''
                 ORDER BY Subunidad) d) AS SubunidadesCorrida,
       (SELECT GROUP_CONCAT(COALESCE(pr.nombreprograma, pr.abreviatura, 'PROG ' || d.IdPrograma), ', ')
          FROM (SELECT DISTINCT IdPrograma FROM {vista} WHERE IdCorrida = @idCorrida ORDER BY IdPrograma) d
          LEFT JOIN programas pr ON pr.id = d.IdPrograma) AS ProgramasCorrida
FROM {vista}
WHERE IdCorrida = @idCorrida;";

                var resumen = await db.QuerySingleAsync<AgStagingResumen>(
                    new CommandDefinition(sqlA, new { idCorrida }, cancellationToken: ct));
                resumen.IdCorrida = idCorrida;

                var parametros = new DynamicParameters();
                string where = BuildWhere(parametros, idCorrida, idSede, idTercero, idPeriodo, idPrograma,
                                          modalidadPrograma, subunidad, semestre, soloViaticas, fechaDesde, fechaHasta);

                string sqlB = $"SELECT COUNT(*) FROM {vista} WHERE 1=1{where}";
                resumen.SesionesConFiltros = await db.ExecuteScalarAsync<int>(
                    new CommandDefinition(sqlB, parametros, cancellationToken: ct));

                return Result.Success(resumen);
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<AgStagingResumen>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[VinculacionQueryRepository] Error en GetResumenStagingAsync", ex);
                return Result.Failure<AgStagingResumen>("Error al construir el resumen de staging: " + ex.Message);
            }
        }


        /// <summary>
        /// Predicado único compartido por el listado de sesiones y los contadores del resumen de staging.
        /// Devuelve un fragmento que empieza por " AND" (o vacío) y acumula los parámetros recibidos.
        /// </summary>
        private static string BuildWhere(
            DynamicParameters parameters,
            long? idCorrida,
            int? idSede,
            int? idTercero,
            int? idPeriodo,
            int? idPrograma,
            string modalidadPrograma,
            string subunidad,
            string semestre,
            bool soloViaticas,
            DateTime? fechaDesde,
            DateTime? fechaHasta)
        {
            var sql = new StringBuilder();

            // Fase 457: filtro por corrida en Modo AG
            if (idCorrida.HasValue && idCorrida.Value > 0)
            {
                sql.Append(" AND IdCorrida = @idCorrida");
                parameters.Add("idCorrida", idCorrida.Value);
            }

            if (idSede.HasValue)
            {
                sql.Append(" AND IdSede = @idSede");
                parameters.Add("idSede", idSede.Value);
            }

            if (idTercero.HasValue)
            {
                sql.Append(" AND IdTercero = @idTercero");
                parameters.Add("idTercero", idTercero.Value);
            }

            if (idPeriodo.HasValue)
            {
                sql.Append(" AND IdPeriodo = @idPeriodo");
                parameters.Add("idPeriodo", idPeriodo.Value);
            }

            if (idPrograma.HasValue)
            {
                sql.Append(" AND IdPrograma = @idPrograma");
                parameters.Add("idPrograma", idPrograma.Value);
            }

            if (!string.IsNullOrEmpty(modalidadPrograma))
            {
                sql.Append(" AND UPPER(TRIM(ProgramaTipo)) = @modalidadPrograma");
                parameters.Add("modalidadPrograma", modalidadPrograma.Trim().ToUpperInvariant());
            }

            if (!string.IsNullOrEmpty(subunidad))
            {
                // Fase 1491: comparación tolerante (mayúsculas/espacios) entre staging y configuración
                sql.Append(" AND UPPER(TRIM(Subunidad)) = UPPER(TRIM(@subunidad))");
                parameters.Add("subunidad", subunidad);
            }

            if (!string.IsNullOrEmpty(semestre))
            {
                sql.Append(" AND Semestre = @semestre");
                parameters.Add("semestre", semestre);
            }

            if (soloViaticas)
            {
                sql.Append(" AND Viatica = 1");
            }

            if (fechaDesde.HasValue)
            {
                sql.Append(" AND date(FechaSesion) >= date(@fechaDesde)");
                parameters.Add("fechaDesde", fechaDesde.Value.ToString("yyyy-MM-dd"));
            }

            if (fechaHasta.HasValue)
            {
                sql.Append(" AND date(FechaSesion) <= date(@fechaHasta)");
                parameters.Add("fechaHasta", fechaHasta.Value.ToString("yyyy-MM-dd"));
            }

            return sql.ToString();
        }

        private static async Task<bool> ExisteTablaAsync(IDbConnection db, string tabla, CancellationToken ct)
        {
            int total = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                "SELECT COUNT(1) FROM sqlite_master WHERE type='table' AND name=@tabla",
                new { tabla },
                cancellationToken: ct));
            return total > 0;
        }

        private static async Task<bool> ExisteVistaAsync(IDbConnection db, string vista, CancellationToken ct)
        {
            int total = await db.ExecuteScalarAsync<int>(new CommandDefinition(
                "SELECT COUNT(1) FROM sqlite_master WHERE type='view' AND name=@vista",
                new { vista },
                cancellationToken: ct));
            return total > 0;
        }

    }
}
