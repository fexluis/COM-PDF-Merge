using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Application.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;
using Dapper;

using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    public class SalvedadConflictoRepository : ISalvedadConflictoRepository
    {
        private readonly IDbConnection _connection;
        private readonly IDbTransaction _transaction;
        private readonly bool _suppressMarkDirty;
        private readonly AgModoDatos _modo;

        public SalvedadConflictoRepository(IDbConnection connection, IDbTransaction transaction = null, bool suppressMarkDirty = false, AgModoDatos modo = null)
        {
            _connection = connection ?? throw new ArgumentNullException(nameof(connection));
            _transaction = transaction;
            _suppressMarkDirty = suppressMarkDirty;
            _modo = modo ?? AgModoDatos.Determinista;
        }

        private void EnsureCrSqliteEnabled()
        {
            var rawConn = _connection;
            if (rawConn is ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context.NonDisposableDbConnection ndc)
            {
                rawConn = ndc.Unwrap();
            }

            if (rawConn is System.Data.SQLite.SQLiteConnection sqliteConn && sqliteConn.State == ConnectionState.Open)
            {
                if (!ConversorPDF.XLL.Core.Sync.CrSqliteLoader.EnableOnConnection(sqliteConn))
                {
                    LoggerHelper.WarningLimitado("crSqliteNoHabilitado", "[SalvedadConflictoRepository] CR-SQLite no pudo habilitarse en la conexión activa. Los cambios no se sincronizarán.", maxVeces: 1);
                }
            }
        }

        private void NotifyMarkDirty()
        {
            if (!_suppressMarkDirty)
            {
                ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();
            }
        }

        public async Task<IReadOnlyCollection<SalvedadConflicto>> GetByPeriodoAsync(int idPeriodo)
        {
            try
            {
                var sql = $@"
                    SELECT 
                        id AS Id,
                        idsesion_a AS IdSesionA,
                        idsesion_b AS IdSesionB,
                        idtercero AS IdTercero,
                        idperiodo AS IdPeriodo,
                        justificacion AS Justificacion,
                        fecha_creacion AS FechaCreacion
                    FROM {_modo.TablaSalvedades}
                    WHERE idperiodo = @idPeriodo;
                ";

                var result = await _connection.QueryAsync<SalvedadConflicto>(sql, new { idPeriodo });
                return result.ToList().AsReadOnly();
            }
            catch (Exception ex)
            {
                // Devolvemos lista vacía en caso de error (e.g., tabla no existe antes de migración)
                LoggerHelper.Warning($"[SalvedadConflictoRepository] Error obteniendo salvedades (probablemente tabla inexistente pre-migración v16): {ex.Message}");
                return new List<SalvedadConflicto>().AsReadOnly();
            }
        }

        public async Task<Result> SaveAsync(SalvedadConflicto salvedad)
        {
            EnsureCrSqliteEnabled();
            try
            {
                var sql = $@"
                    INSERT OR IGNORE INTO {_modo.TablaSalvedades} 
                        (idsesion_a, idsesion_b, idtercero, idperiodo, justificacion, fecha_creacion)
                    VALUES 
                        (@IdSesionA, @IdSesionB, @IdTercero, @IdPeriodo, @Justificacion, @FechaCreacion);
                ";

                // Asegurar que A es menor que B para unicidad
                int idA = Math.Min(salvedad.IdSesionA, salvedad.IdSesionB);
                int idB = Math.Max(salvedad.IdSesionA, salvedad.IdSesionB);

                await _connection.ExecuteAsync(sql, new
                {
                    IdSesionA = idA,
                    IdSesionB = idB,
                    salvedad.IdTercero,
                    salvedad.IdPeriodo,
                    salvedad.Justificacion,
                    FechaCreacion = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                });

                NotifyMarkDirty();
                return Result.Success();
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al guardar salvedad: {ex.Message}");
            }
        }

        public async Task<Result> DeleteAsync(int idSesionA, int idSesionB)
        {
            EnsureCrSqliteEnabled();
            try
            {
                int idA = Math.Min(idSesionA, idSesionB);
                int idB = Math.Max(idSesionA, idSesionB);

                var sql = $"DELETE FROM {_modo.TablaSalvedades} WHERE idsesion_a = @idA AND idsesion_b = @idB;";
                await _connection.ExecuteAsync(sql, new { idA, idB });

                NotifyMarkDirty();
                return Result.Success();
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al eliminar salvedad: {ex.Message}");
            }
        }
    }
}
