using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using CSharpFunctionalExtensions;
using Dapper;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    public class CdpDependenciaRepository
    {
        private readonly IDbConnection _connection;

        public CdpDependenciaRepository(IDbConnection connection)
        {
            _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        }

        private const string SqlSelectAll = @"
            SELECT id AS Id, idperiodo AS IdPeriodo, cdpnum AS CdpNum, fechacdp AS FechaCdp,
                   rubro AS Rubro, responsable AS Responsable, cargo AS Cargo,
                   elaboro AS Elaboro, reviso AS Reviso, dependencia AS Dependencia,
                   dependencia_full AS DependenciaFull, activo AS Activo, subunidad AS Subunidad
            FROM cdpDependencias";

        public Result<IReadOnlyCollection<CdpDependencia>> GetAll()
        {
            try
            {
                var subunidad = ConversorPDF.XLL.Core.Sync.SyncContext.ActiveSubunidad;
                var query = SqlSelectAll + " WHERE (subunidad = @Subunidad OR subunidad = '' OR subunidad IS NULL) ORDER BY id;";
                var result = _connection.Query<CdpDependencia>(query, new { Subunidad = subunidad }).ToList();
                return Result.Success((IReadOnlyCollection<CdpDependencia>)result);
            }
            catch (Exception ex)
            {
                return Result.Failure<IReadOnlyCollection<CdpDependencia>>($"Error al obtener cdpDependencias: {ex.Message}");
            }
        }

        public Result<CdpDependencia> GetById(int id)
        {
            try
            {
                var sql = SqlSelectAll + " WHERE id = @Id;";
                var result = _connection.QueryFirstOrDefault<CdpDependencia>(sql, new { Id = id });
                if (result == null)
                    return Result.Failure<CdpDependencia>("No se encontró el registro.");
                return Result.Success(result);
            }
            catch (Exception ex)
            {
                return Result.Failure<CdpDependencia>($"Error al obtener cdpDependencia: {ex.Message}");
            }
        }

        public Result<int> Add(CdpDependencia cdp)
        {
            try
            {
                var sql = @"
                    INSERT INTO cdpDependencias (idperiodo, cdpnum, fechacdp, rubro, responsable, cargo, elaboro, reviso, dependencia, dependencia_full, activo, subunidad) 
                    VALUES (@IdPeriodo, @CdpNum, @FechaCdp, @Rubro, @Responsable, @Cargo, @Elaboro, @Reviso, @Dependencia, @DependenciaFull, @Activo, @Subunidad);
                    SELECT last_insert_rowid();";
                
                var id = _connection.ExecuteScalar<int>(sql, cdp);
                ConversorPDF.XLL.Core.Repositories.ViaticosRepository.InvalidarCacheDependenciaFull();
                return Result.Success(id);
            }
            catch (Exception ex)
            {
                return Result.Failure<int>($"Error al agregar cdpDependencia: {ex.Message}");
            }
        }

        public Result Update(CdpDependencia cdp)
        {
            try
            {
                var sql = @"
                    UPDATE cdpDependencias SET 
                        idperiodo = @IdPeriodo,
                        cdpnum = @CdpNum,
                        fechacdp = @FechaCdp,
                        rubro = @Rubro,
                        responsable = @Responsable,
                        cargo = @Cargo,
                        elaboro = @Elaboro,
                        reviso = @Reviso,
                        dependencia = @Dependencia,
                        dependencia_full = @DependenciaFull,
                        activo = @Activo,
                        subunidad = @Subunidad
                    WHERE id = @Id;";
                
                var affected = _connection.Execute(sql, cdp);
                if (affected == 0) return Result.Failure("No se encontró el registro para actualizar.");
                ConversorPDF.XLL.Core.Repositories.ViaticosRepository.InvalidarCacheDependenciaFull();
                return Result.Success();
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al actualizar cdpDependencia: {ex.Message}");
            }
        }

        public Result Delete(int id)
        {
            try
            {
                var sql = "DELETE FROM cdpDependencias WHERE id = @Id;";
                var affected = _connection.Execute(sql, new { Id = id });
                if (affected == 0) return Result.Failure("No se encontró el registro para eliminar.");
                ConversorPDF.XLL.Core.Repositories.ViaticosRepository.InvalidarCacheDependenciaFull();
                return Result.Success();
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al eliminar cdpDependencia: {ex.Message}");
            }
        }
    }
}
