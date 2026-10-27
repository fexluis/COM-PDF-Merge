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
    public class ProgramaLookupRepository : IProgramaLookupRepository
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public ProgramaLookupRepository(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        public async Task<Result<IReadOnlyCollection<LookupItem>>> GetProgramasLookupAsync(CancellationToken ct)
        {
            try
            {
                var sql = "SELECT id as Id, COALESCE(abreviatura, nombreprograma) as DisplayName FROM programas ORDER BY COALESCE(abreviatura, nombreprograma)";
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
                LoggerHelper.Warning("[ProgramaLookupRepository] Conexión dispuesta, reintentando...");
                try
                {
                    using (var freshCtx = SqliteHelper.GetContextForPath())
                    {
                        var newDb = freshCtx.Open();
                        var sql = "SELECT id as Id, COALESCE(abreviatura, nombreprograma) as DisplayName FROM programas ORDER BY COALESCE(abreviatura, nombreprograma)";
                        var items = await newDb.QueryAsync<LookupItem>(new CommandDefinition(sql, cancellationToken: ct));
                        return Result.Success<IReadOnlyCollection<LookupItem>>(items.ToList().AsReadOnly());
                    }
                }
                catch (Exception exRetry)
                {
                    LoggerHelper.Error("[ProgramaLookupRepository] Error al reintentar obtener programas", exRetry);
                    return Result.Failure<IReadOnlyCollection<LookupItem>>($"Error: {exRetry.Message}");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[ProgramaLookupRepository] Error al obtener programas", ex);
                return Result.Failure<IReadOnlyCollection<LookupItem>>($"Error: {ex.Message}");
            }
        }
    }
}
