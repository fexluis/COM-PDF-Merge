using System;
using System.Data;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Repositories
{
    public class ConfigRepository : IConfigRepository
    {
        private readonly EncryptedSQLiteContext _context;
        private readonly IDbConnection _connection;

        /// <summary>
        /// Inicializa el repositorio con el contexto encriptado de SQLite.
        /// </summary>
        public ConfigRepository(EncryptedSQLiteContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Inicializa el repositorio con una conexión existente.
        /// </summary>
        public ConfigRepository(IDbConnection connection)
        {
            _connection = connection;
        }

        private IDbConnection GetConnection() => _connection ?? _context.Open();

        public Result<Config> GetConfig()
        {
            try
            {
                var db = GetConnection();
                const string sql = @"SELECT id as Id, subunidad as Subunidad, id_territorial as IdTerritorial,
                    id_periodo as IdPeriodo, fecha_inicio as FechaInicio, fecha_fin as FechaFin, cdp as Cdp,
                    salario_decreto_613_2025 as SalarioDecreto6132025,
                    viatico_diario_tarifa_a as ViaticoDiarioTarifaA,
                    viatico_diario_tarifa_b as ViaticoDiarioTarifaB,
                    porcentaje_decreto_a as PorcentajeDecretoA,
                    porcentaje_decreto_b as PorcentajeDecretoB,
                    dt_nombre as DtNombre, dt_cargo as DtCargo,
                    corporativo_cedula as CorporativoCedula,
                    corporativo_nombre as CorporativoNombre,
                    corporativo_cargo as CorporativoCargo,
                    departamento_base as DepartamentoBase, ciudad_base as CiudadBase,
                    aeropuerto_base as AeropuertoBase, created_at as CreatedAt, updated_at as UpdatedAt
                    FROM config WHERE id = 1";
                var config = db.QuerySingleOrDefault<Config>(sql);
                if (config == null)
                    return Result.Failure<Config>("No se encontró la configuración (id=1).");
                return Result.Success(config);
            }
            catch (Exception ex)
            {
                return Result.Failure<Config>($"Error al obtener la configuración: {ex.Message}");
            }
        }

        public Result UpdateConfig(Config config)
        {
            try
            {
                var db = GetConnection();
                const string sql = @"UPDATE config SET 
                    subunidad = @Subunidad, id_territorial = @IdTerritorial, id_periodo = @IdPeriodo,
                    fecha_inicio = @FechaInicio, fecha_fin = @FechaFin, cdp = @Cdp,
                    salario_decreto_613_2025 = @SalarioDecreto6132025,
                    viatico_diario_tarifa_a = @ViaticoDiarioTarifaA,
                    viatico_diario_tarifa_b = @ViaticoDiarioTarifaB,
                    porcentaje_decreto_a = @PorcentajeDecretoA,
                    porcentaje_decreto_b = @PorcentajeDecretoB,
                    dt_nombre = @DtNombre, dt_cargo = @DtCargo,
                    corporativo_cedula = @CorporativoCedula,
                    corporativo_nombre = @CorporativoNombre,
                    corporativo_cargo = @CorporativoCargo,
                    departamento_base = @DepartamentoBase, ciudad_base = @CiudadBase,
                    aeropuerto_base = @AeropuertoBase, updated_at = CURRENT_TIMESTAMP
                    WHERE id = 1";
                int rows = db.Execute(sql, config);
                if (rows == 0)
                    return Result.Failure("No se pudo actualizar la configuración.");
                
                // Cambio (Fase 7): Notificar cambios al Vault para encriptación al cerrar
                _context?.MarkChangesPersisted();
                return Result.Success();
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al actualizar la configuración: {ex.Message}");
            }
        }

        public Result UpdateSubunidad(string subunidad)
        {
            try
            {
                var db = GetConnection();
                const string sql = "UPDATE config SET subunidad = @Subunidad, updated_at = CURRENT_TIMESTAMP WHERE id = 1";
                int rows = db.Execute(sql, new { Subunidad = (subunidad ?? string.Empty).Trim().ToUpperInvariant() });
                if (rows == 0)
                    return Result.Failure("No se pudo actualizar la subunidad en la configuración.");

                _context?.MarkChangesPersisted();
                return Result.Success();
            }
            catch (Exception ex)
            {
                return Result.Failure($"Error al actualizar la subunidad: {ex.Message}");
            }
        }
    }
}
