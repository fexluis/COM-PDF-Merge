using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CSharpFunctionalExtensions;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    /// <summary>
    /// Repositorio para persistencia de selecciones de filtros y terceros en ag_filtro_seleccion (Fase 1315 / 1348 / D10).
    /// </summary>
    public class AgFiltroSeleccionRepository : IAgFiltroSeleccionRepository
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public AgFiltroSeleccionRepository(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        // Fase 1441: Normalización simétrica de subunidad en lectura y escritura para evitar desalineaciones por espacios o nulos
        private static string NormalizarSub(string s) => (s ?? string.Empty).Trim();

        public Result<List<int>> Obtener(string subunidad, string tipo)
        {
            try
            {
                var sub = NormalizarSub(subunidad);
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    // Fase 1429 & 1441: Evitar unión accidental de todas las subunidades cuando @s es vacío
                    const string sql = @"
SELECT refid
FROM ag_filtro_seleccion
WHERE tipo = @t
  AND (subunidad = @s OR (@s <> '' AND subunidad = ''))
ORDER BY refid;";

                    var ids = db.Query<int>(sql, new { s = sub, t = tipo ?? "" }).ToList();
                    return Result.Success(ids);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFiltroSeleccionRepository] Error al obtener filtros: " + ex.Message);
                return Result.Failure<List<int>>(ex.Message);
            }
        }

        public async Task<Result<List<int>>> ObtenerAsync(string subunidad, string tipo, CancellationToken ct = default)
        {
            try
            {
                var sub = NormalizarSub(subunidad);
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();
                    // Fase 1429 & 1441: Evitar unión accidental de todas las subunidades cuando @s es vacío
                    const string sql = @"
SELECT refid
FROM ag_filtro_seleccion
WHERE tipo = @t
  AND (subunidad = @s OR (@s <> '' AND subunidad = ''))
ORDER BY refid;";

                    var ids = (await db.QueryAsync<int>(new CommandDefinition(sql, new { s = sub, t = tipo ?? "" }, cancellationToken: ct))).ToList();
                    return Result.Success(ids);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFiltroSeleccionRepository] Error asíncrono al obtener filtros: " + ex.Message);
                return Result.Failure<List<int>>(ex.Message);
            }
        }

        public Result Reemplazar(string subunidad, string tipo, IEnumerable<int> refIds, AgDimensionesSeleccion dims = null)
        {
            try
            {
                var sub = NormalizarSub(subunidad);
                var idsList = refIds?.ToList() ?? new List<int>();
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    using (var tx = db.BeginTransaction())
                    {
                        db.Execute(
                            "DELETE FROM ag_filtro_seleccion WHERE subunidad = @s AND tipo = @t;",
                            new { s = sub, t = tipo ?? "" },
                            transaction: tx);

                        if (idsList.Count > 0)
                        {
                            const string insertSql = @"
INSERT OR IGNORE INTO ag_filtro_seleccion(subunidad, tipo, refid, semestre, grupo, modalidad)
VALUES (@s, @t, @r, @semestre, @grupo, @modalidad);";

                            var paramsList = idsList.Select(r => new
                            {
                                s = sub,
                                t = tipo ?? "",
                                r = r,
                                semestre = dims?.Semestre ?? "",
                                grupo = dims?.Grupo ?? "",
                                modalidad = dims?.Modalidad ?? ""
                            });

                            db.Execute(insertSql, paramsList, transaction: tx);
                        }

                        tx.Commit();
                    }

                    return Result.Success();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFiltroSeleccionRepository] Error al reemplazar filtros: " + ex.Message);
                return Result.Failure(ex.Message);
            }
        }

        public async Task<Result> ReemplazarAsync(string subunidad, string tipo, IEnumerable<int> refIds, AgDimensionesSeleccion dims = null, CancellationToken ct = default)
        {
            try
            {
                var sub = NormalizarSub(subunidad);
                var idsList = refIds?.ToList() ?? new List<int>();
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    using (var tx = db.BeginTransaction())
                    {
                        await db.ExecuteAsync(new CommandDefinition(
                            "DELETE FROM ag_filtro_seleccion WHERE subunidad = @s AND tipo = @t;",
                            new { s = sub, t = tipo ?? "" },
                            transaction: tx,
                            cancellationToken: ct));

                        if (idsList.Count > 0)
                        {
                            const string insertSql = @"
INSERT OR IGNORE INTO ag_filtro_seleccion(subunidad, tipo, refid, semestre, grupo, modalidad)
VALUES (@s, @t, @r, @semestre, @grupo, @modalidad);";

                            var paramsList = idsList.Select(r => new
                            {
                                s = sub,
                                t = tipo ?? "",
                                r = r,
                                semestre = dims?.Semestre ?? "",
                                grupo = dims?.Grupo ?? "",
                                modalidad = dims?.Modalidad ?? ""
                            });

                            await db.ExecuteAsync(new CommandDefinition(insertSql, paramsList, transaction: tx, cancellationToken: ct));
                        }

                        tx.Commit();
                    }

                    return Result.Success();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFiltroSeleccionRepository] Error asíncrono al reemplazar filtros: " + ex.Message);
                return Result.Failure(ex.Message);
            }
        }

        public Result Limpiar(string subunidad, string tipo)
        {
            try
            {
                var sub = NormalizarSub(subunidad);
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    db.Execute(
                        "DELETE FROM ag_filtro_seleccion WHERE subunidad = @s AND tipo = @t;",
                        new { s = sub, t = tipo ?? "" });
                    return Result.Success();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFiltroSeleccionRepository] Error al limpiar filtros: " + ex.Message);
                return Result.Failure(ex.Message);
            }
        }

        public async Task<Result> LimpiarAsync(string subunidad, string tipo, CancellationToken ct = default)
        {
            try
            {
                var sub = NormalizarSub(subunidad);
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();
                    await db.ExecuteAsync(new CommandDefinition(
                        "DELETE FROM ag_filtro_seleccion WHERE subunidad = @s AND tipo = @t;",
                        new { s = sub, t = tipo ?? "" },
                        cancellationToken: ct));
                    return Result.Success();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFiltroSeleccionRepository] Error asíncrono al limpiar filtros: " + ex.Message);
                return Result.Failure(ex.Message);
            }
        }

        public async Task<Result<AgFiltrosSeleccion>> ObtenerFiltrosAsync(string subunidad, CancellationToken ct = default)
        {
            try
            {
                var f = new AgFiltrosSeleccion();
                var perRes = await ObtenerAsync(subunidad, AgFiltroTipo.Periodo, ct);
                if (perRes.IsFailure) return Result.Failure<AgFiltrosSeleccion>("Periodo: " + perRes.Error);
                f.PeriodosIds = perRes.Value;

                var sedRes = await ObtenerAsync(subunidad, AgFiltroTipo.Sede, ct);
                if (sedRes.IsFailure) return Result.Failure<AgFiltrosSeleccion>("Sede: " + sedRes.Error);
                f.SedesIds = sedRes.Value;

                var progRes = await ObtenerAsync(subunidad, AgFiltroTipo.Programa, ct);
                if (progRes.IsFailure) return Result.Failure<AgFiltrosSeleccion>("Programa: " + progRes.Error);
                f.ProgramasIds = progRes.Value;

                var terRes = await ObtenerAsync(subunidad, AgFiltroTipo.Tercero, ct);
                if (terRes.IsFailure) return Result.Failure<AgFiltrosSeleccion>("Tercero: " + terRes.Error);
                f.TercerosIds = terRes.Value;

                var tipoRes = await ObtenerAsync(subunidad, AgFiltroTipo.TipoPrograma, ct);
                if (tipoRes.IsFailure) return Result.Failure<AgFiltrosSeleccion>("TipoPrograma: " + tipoRes.Error);
                if (tipoRes.Value.Count > 0)
                {
                    int refId = tipoRes.Value[0];
                    if (refId == 1) f.TipoPrograma = AgTipoProgramaCatalog.Pregrado;
                    else if (refId == 2) f.TipoPrograma = AgTipoProgramaCatalog.Posgrado;
                    else f.TipoPrograma = AgTipoProgramaCatalog.Todos;
                }

                return Result.Success(f);
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFiltroSeleccionRepository] Error en ObtenerFiltrosAsync: " + ex.Message);
                return Result.Failure<AgFiltrosSeleccion>(ex.Message);
            }
        }

        public async Task<Result> GuardarFiltrosAsync(AgFiltrosSeleccion filtros, string subunidad, CancellationToken ct = default)
        {
            if (filtros == null) return Result.Failure("Filtros nulos");
            try
            {
                var r1 = await ReemplazarAsync(subunidad, AgFiltroTipo.Periodo, filtros.PeriodosIds, null, ct);
                if (r1.IsFailure) return r1;

                var r2 = await ReemplazarAsync(subunidad, AgFiltroTipo.Sede, filtros.SedesIds, null, ct);
                if (r2.IsFailure) return r2;

                var r3 = await ReemplazarAsync(subunidad, AgFiltroTipo.Programa, filtros.ProgramasIds, null, ct);
                if (r3.IsFailure) return r3;

                var r4 = await ReemplazarAsync(subunidad, AgFiltroTipo.Tercero, filtros.TercerosIds, null, ct);
                if (r4.IsFailure) return r4;

                int tipoRef = 0;
                if (string.Equals(filtros.TipoPrograma, AgTipoProgramaCatalog.Pregrado, StringComparison.OrdinalIgnoreCase)) tipoRef = 1;
                else if (string.Equals(filtros.TipoPrograma, AgTipoProgramaCatalog.Posgrado, StringComparison.OrdinalIgnoreCase)) tipoRef = 2;

                var r5 = await ReemplazarAsync(subunidad, AgFiltroTipo.TipoPrograma, new[] { tipoRef }, null, ct);
                if (r5.IsFailure) return r5;

                return Result.Success();
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgFiltroSeleccionRepository] Error en GuardarFiltrosAsync: " + ex.Message);
                return Result.Failure(ex.Message);
            }
        }
    }
}
