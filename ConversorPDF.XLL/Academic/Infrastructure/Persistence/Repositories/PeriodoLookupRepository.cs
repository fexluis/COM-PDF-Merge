using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Academic.Domain.Models;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    public class PeriodoLookupRepository : IPeriodoLookupRepository
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public PeriodoLookupRepository(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        public async Task<Result<IReadOnlyCollection<LookupItem>>> GetPeriodosLookupAsync(CancellationToken ct)
        {
            try
            {
                var sql = "SELECT idperiodo as Id, periodo as DisplayName FROM periodos ORDER BY idperiodo DESC";
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();
                    var items = await db.QueryAsync<LookupItem>(new CommandDefinition(sql, cancellationToken: ct));
                    return Result.Success<IReadOnlyCollection<LookupItem>>(items.ToList().AsReadOnly());
                }
            }
            catch (ObjectDisposedException)
            {
                LoggerHelper.Warning("[PeriodoLookupRepository] Conexión dispuesta, reintentando...");
                try
                {
                    using (var freshCtx = SqliteHelper.GetContextForPath())
                    {
                        var newDb = freshCtx.Open();
                        var sql = "SELECT idperiodo as Id, periodo as DisplayName FROM periodos ORDER BY idperiodo DESC";
                        var items = await newDb.QueryAsync<LookupItem>(new CommandDefinition(sql, cancellationToken: ct));
                        return Result.Success<IReadOnlyCollection<LookupItem>>(items.ToList().AsReadOnly());
                    }
                }
                catch (Exception exRetry)
                {
                    LoggerHelper.Error("[PeriodoLookupRepository] Error al reintentar obtener periodos", exRetry);
                    return Result.Failure<IReadOnlyCollection<LookupItem>>($"Error: {exRetry.Message}");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[PeriodoLookupRepository] Error al obtener periodos", ex);
                return Result.Failure<IReadOnlyCollection<LookupItem>>($"Error: {ex.Message}");
            }
        }
    }
}
