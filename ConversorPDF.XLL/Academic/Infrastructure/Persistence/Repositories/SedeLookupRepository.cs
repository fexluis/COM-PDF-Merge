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
    public class SedeLookupRepository : ISedeLookupRepository
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public SedeLookupRepository(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        public async Task<Result<IReadOnlyCollection<LookupItem>>> GetSedesLookupAsync(CancellationToken ct)
        {
            try
            {
                var sql = "SELECT id as Id, CASE WHEN subsede IS NOT NULL AND subsede != '' THEN ciudad || ' (' || subsede || ')' ELSE ciudad END as DisplayName, subunidad as Subunidad FROM sedes WHERE activo = 1 ORDER BY ciudad, subsede";
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
                // El contexto fue dispuesto mientras consultábamos - intentar una vez más
                LoggerHelper.Warning("[SedeLookupRepository] Conexión dispuesta, reintentando...");
                try
                {
                    using (var freshCtx = SqliteHelper.GetContextForPath())
                    {
                        var newDb = freshCtx.Open();
                        var sql = "SELECT id as Id, CASE WHEN subsede IS NOT NULL AND subsede != '' THEN ciudad || ' (' || subsede || ')' ELSE ciudad END as DisplayName, subunidad as Subunidad FROM sedes WHERE activo = 1 ORDER BY ciudad, subsede";
                        var items = await newDb.QueryAsync<LookupItem>(new CommandDefinition(sql, cancellationToken: ct));
                        return Result.Success<IReadOnlyCollection<LookupItem>>(items.ToList().AsReadOnly());
                    }
                }
                catch (Exception exRetry)
                {
                    LoggerHelper.Error("[SedeLookupRepository] Error al reintentar obtener sedes para lookup", exRetry);
                    return Result.Failure<IReadOnlyCollection<LookupItem>>($"Error: {exRetry.Message}");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[SedeLookupRepository] Error al obtener sedes para lookup", ex);
                return Result.Failure<IReadOnlyCollection<LookupItem>>($"Error: {ex.Message}");
            }
        }
    }
}
