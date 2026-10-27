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
    public class TerceroLookupRepository : ITerceroLookupRepository
    {
        private const string LookupSql = @"
SELECT 
    id as Id, 
    cedula as Cedula,
    CASE 
        WHEN residencia IS NULL OR trim(residencia) = '' THEN nombre 
        ELSE nombre || ' (' || residencia || ')' 
    END as DisplayName, 
    subunidad as Subunidad 
FROM terceros 
WHERE (cedula IS NULL OR cedula <> '1111111111')
  AND (nombre IS NULL OR UPPER(TRIM(nombre)) <> 'VACANTE')
ORDER BY nombre";

        private readonly Func<IDbConnection> _connectionFactory;

        public TerceroLookupRepository(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        public async Task<Result<IReadOnlyCollection<LookupItem>>> GetTercerosLookupAsync(CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();
                    var items = await db.QueryAsync<LookupItem>(new CommandDefinition(LookupSql, cancellationToken: ct));
                    return Result.Success<IReadOnlyCollection<LookupItem>>(items.ToList().AsReadOnly());
                }
            }
            catch (ObjectDisposedException)
            {
                LoggerHelper.Warning("[TerceroLookupRepository] Conexión dispuesta, reintentando...");
                try
                {
                    using (var freshCtx = SqliteHelper.GetContextForPath())
                    {
                        var newDb = freshCtx.Open();
                        var items = await newDb.QueryAsync<LookupItem>(new CommandDefinition(LookupSql, cancellationToken: ct));
                        return Result.Success<IReadOnlyCollection<LookupItem>>(items.ToList().AsReadOnly());
                    }
                }
                catch (Exception exRetry)
                {
                    LoggerHelper.Error("[TerceroLookupRepository] Error al reintentar obtener terceros para lookup", exRetry);
                    return Result.Failure<IReadOnlyCollection<LookupItem>>($"Error: {exRetry.Message}");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[TerceroLookupRepository] Error al obtener terceros para lookup", ex);
                return Result.Failure<IReadOnlyCollection<LookupItem>>($"Error: {ex.Message}");
            }
        }
    }
}
