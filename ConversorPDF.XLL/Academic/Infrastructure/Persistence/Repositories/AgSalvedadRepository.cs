using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>
    /// Repositorio para la gestión de salvedades de traslados/conflictos en la capa de staging del AG.
    /// </summary>
    public class AgSalvedadRepository : IAgSalvedadRepository
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public AgSalvedadRepository(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        public async Task<Result<IReadOnlyCollection<AgSalvedad>>> GetByCorridaAsync(long idCorrida, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string sql = "SELECT id as Id, id_corrida as IdCorrida, idsesiona as IdSesionA, idsesionb as IdSesionB, idtercero as IdTercero, idperiodo as IdPeriodo, justificacion as Justificacion, subunidad as Subunidad FROM ag_salvedades_conflictos WHERE id_corrida = @idCorrida";
                    var items = (await db.QueryAsync<AgSalvedad>(new CommandDefinition(sql, new { idCorrida }, cancellationToken: ct))).ToList();
                    return Result.Success<IReadOnlyCollection<AgSalvedad>>(items.AsReadOnly());
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<IReadOnlyCollection<AgSalvedad>>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[AgSalvedadRepository] Error al obtener salvedades para corrida {idCorrida}", ex);
                return Result.Failure<IReadOnlyCollection<AgSalvedad>>($"Error al leer salvedades: {ex.Message}");
            }
        }

        public async Task<Result> SaveAsync(AgSalvedad salvedad, CancellationToken ct)
        {
            if (salvedad == null) return Result.Failure("Salvedad nula");
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string sql = @"
INSERT OR REPLACE INTO ag_salvedades_conflictos (id_corrida, idsesiona, idsesionb, idtercero, idperiodo, justificacion, subunidad)
VALUES (@IdCorrida, @IdSesionA, @IdSesionB, @IdTercero, @IdPeriodo, @Justificacion, @Subunidad);";

                    await db.ExecuteAsync(new CommandDefinition(sql, salvedad, cancellationToken: ct));
                    return Result.Success();
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgSalvedadRepository] Error guardando salvedad", ex);
                return Result.Failure($"Error al guardar salvedad: {ex.Message}");
            }
        }

        public async Task<Result> DeleteAsync(long idCorrida, int idSesionA, int idSesionB, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    string sql = "DELETE FROM ag_salvedades_conflictos WHERE id_corrida = @idCorrida AND idsesiona = @idSesionA AND idsesionb = @idSesionB";
                    await db.ExecuteAsync(new CommandDefinition(sql, new { idCorrida, idSesionA, idSesionB }, cancellationToken: ct));
                    return Result.Success();
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgSalvedadRepository] Error eliminando salvedad", ex);
                return Result.Failure($"Error al eliminar salvedad: {ex.Message}");
            }
        }
    }
}
