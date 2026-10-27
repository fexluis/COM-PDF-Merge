using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;
using CSharpFunctionalExtensions;
using Dapper;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Implementación de la resolución de ventana semestral (α) para el motor genético.
    /// Resuelve en cascada: Festivos(PRESENCIAL) -> Manual -> Error bloqueante.
    /// </summary>
    public class AgResolutorVentanaSemestre : IAgResolutorVentanaSemestre
    {
        private readonly Func<IDbConnection> _connectionFactory;

        public AgResolutorVentanaSemestre(Func<IDbConnection> connectionFactory)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
        }

        public async Task<Result<VentanaSemestre>> ResolverVentanaAsync(int idPeriodo, DateTime? fechaInicioManual, DateTime? fechaFinManual, CancellationToken ct)
        {
            try
            {
                using (var db = _connectionFactory())
                {
                    if (db.State != ConnectionState.Open) db.Open();
                    ct.ThrowIfCancellationRequested();

                    // 1. Obtener información del periodo verificando dinámicamente columnas de ventana (Fase 185)
                    bool tieneColumnasVentana = PeriodosTieneVentana(db);
                    string colInicio = tieneColumnasVentana ? "periodo_inicio" : "NULL";
                    string colFin = tieneColumnasVentana ? "periodo_fin" : "NULL";

                    string sqlPeriodo = idPeriodo > 0
                        ? $"SELECT idperiodo as Id, periodo as Nombre, activo as Activo, {colInicio} as PeriodoInicio, {colFin} as PeriodoFin FROM periodos WHERE idperiodo = @idPeriodo LIMIT 1"
                        : $"SELECT idperiodo as Id, periodo as Nombre, activo as Activo, {colInicio} as PeriodoInicio, {colFin} as PeriodoFin FROM periodos WHERE activo = 1 ORDER BY idperiodo DESC LIMIT 1";

                    var periodoInfo = await db.QuerySingleOrDefaultAsync(new CommandDefinition(sqlPeriodo, new { idPeriodo }, cancellationToken: ct));
                    if (periodoInfo == null)
                    {
                        return Result.Failure<VentanaSemestre>("No se encontró ningún periodo académico activo o correspondiente al ID suministrado.");
                    }

                    int resolvedIdPeriodo = (int)periodoInfo.Id;
                    string nombrePeriodo = (string)periodoInfo.Nombre ?? string.Empty;
                    string perIniRaw = periodoInfo.PeriodoInicio?.ToString();
                    string perFinRaw = periodoInfo.PeriodoFin?.ToString();

                    // Prioridad 1: Fechas canónicas en la tabla 'periodos' (Fase 154)
                    if (TryParseFecha(perIniRaw, out DateTime perIni) && TryParseFecha(perFinRaw, out DateTime perFin) && perFin.Date >= perIni.Date)
                    {
                        if (fechaInicioManual.HasValue && fechaFinManual.HasValue && fechaInicioManual.Value <= fechaFinManual.Value)
                        {
                            var manualIni = fechaInicioManual.Value.Date;
                            var manualFin = fechaFinManual.Value.Date;

                            return VentanaSemestre.Crear(
                                resolvedIdPeriodo,
                                nombrePeriodo,
                                manualIni,
                                manualFin,
                                AgOrigenVentana.Manual,
                                $"Rango manual acotado dentro del periodo {nombrePeriodo} (Base periodo: {perIni:yyyy-MM-dd} a {perFin:yyyy-MM-dd})");
                        }

                        return VentanaSemestre.Crear(
                            resolvedIdPeriodo,
                            nombrePeriodo,
                            perIni.Date,
                            perFin.Date,
                            AgOrigenVentana.Periodos,
                            $"Ventana resuelta desde la configuración canónica de la tabla 'periodos' ({perIni:yyyy-MM-dd} a {perFin:yyyy-MM-dd}).");
                    }

                    // Prioridad 2: Intentar resolver desde festivos (tipofestivo = 'PRESENCIAL')
                    string sqlPresenciales = @"
SELECT fechafestivoinicio as InicioRaw, fechafestivofinal as FinRaw
FROM festivos
WHERE idperiodo = @resolvedIdPeriodo AND UPPER(TRIM(tipofestivo)) = 'PRESENCIAL'";

                    var filasFestivos = (await db.QueryAsync(new CommandDefinition(sqlPresenciales, new { resolvedIdPeriodo }, cancellationToken: ct))).ToList();

                    DateTime? minPresencial = null;
                    DateTime? maxPresencial = null;

                    foreach (var f in filasFestivos)
                    {
                        string iniStr = f.InicioRaw?.ToString();
                        string finStr = f.FinRaw?.ToString();

                        if (TryParseFecha(iniStr, out DateTime dIni))
                        {
                            if (!minPresencial.HasValue || dIni < minPresencial.Value) minPresencial = dIni;
                            if (!maxPresencial.HasValue || dIni > maxPresencial.Value) maxPresencial = dIni;
                        }

                        if (TryParseFecha(finStr, out DateTime dFin))
                        {
                            if (!minPresencial.HasValue || dFin < minPresencial.Value) minPresencial = dFin;
                            if (!maxPresencial.HasValue || dFin > maxPresencial.Value) maxPresencial = dFin;
                        }
                    }

                    if (minPresencial.HasValue && maxPresencial.HasValue && minPresencial.Value <= maxPresencial.Value)
                    {
                        if (fechaInicioManual.HasValue && fechaFinManual.HasValue && fechaInicioManual.Value <= fechaFinManual.Value)
                        {
                            var manualIni = fechaInicioManual.Value.Date;
                            var manualFin = fechaFinManual.Value.Date;

                            return VentanaSemestre.Crear(
                                resolvedIdPeriodo,
                                nombrePeriodo,
                                manualIni,
                                manualFin,
                                AgOrigenVentana.Manual,
                                $"Rango manual acotado dentro del periodo {nombrePeriodo} (Presencial base: {minPresencial.Value:yyyy-MM-dd} a {maxPresencial.Value:yyyy-MM-dd})");
                        }

                        return VentanaSemestre.Crear(
                            resolvedIdPeriodo,
                            nombrePeriodo,
                            minPresencial.Value.Date,
                            maxPresencial.Value.Date,
                            AgOrigenVentana.Presencial,
                            $"Ventana derivada de {filasFestivos.Count} fechas presenciales configuradas en el periodo.");
                    }

                    // Prioridad 3: Revisar si el usuario especificó rango manual válido
                    if (fechaInicioManual.HasValue && fechaFinManual.HasValue && fechaInicioManual.Value <= fechaFinManual.Value)
                    {
                        return VentanaSemestre.Crear(
                            resolvedIdPeriodo,
                            nombrePeriodo,
                            fechaInicioManual.Value.Date,
                            fechaFinManual.Value.Date,
                            AgOrigenVentana.Manual,
                            $"Ventana manual definida por el usuario para el periodo {nombrePeriodo}.");
                    }

                    // Prioridad 4: Bloqueo si no hay fechas válidas
                    return Result.Failure<VentanaSemestre>(
                        $"No se pudo determinar la ventana temporal para el periodo '{nombrePeriodo}' (Id={resolvedIdPeriodo}). " +
                        "Debe configurar las fechas de inicio y fin en la tabla 'periodos', al menos un rango 'PRESENCIAL' en Festivos o ingresar las fechas manualmente.");
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Failure<VentanaSemestre>("Operación cancelada.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[AgResolutorVentanaSemestre] Error resolviendo ventana del semestre", ex);
                return Result.Failure<VentanaSemestre>($"Error al resolver ventana del semestre: {ex.Message}");
            }
        }

        public Result<VentanaSemestre> ResolverVentana(int idPeriodo, DateTime? fechaInicioManual, DateTime? fechaFinManual)
        {
            return ResolverVentanaAsync(idPeriodo, fechaInicioManual, fechaFinManual, CancellationToken.None).GetAwaiter().GetResult();
        }

        private static bool TryParseFecha(string raw, out DateTime result)
        {
            result = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(raw)) return false;

            string[] formats = { "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "yyyy/MM/dd", "yyyyMMdd" };
            if (DateTime.TryParseExact(raw.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
            {
                return true;
            }

            return DateTime.TryParse(raw.Trim(), out result);
        }

        private static bool PeriodosTieneVentana(IDbConnection db)
        {
            try
            {
                var cols = db.Query<string>("SELECT name FROM pragma_table_info('periodos');")
                             .Select(x => x.ToLowerInvariant())
                             .ToList();
                return cols.Contains("periodo_inicio") && cols.Contains("periodo_fin");
            }
            catch
            {
                return false;
            }
        }
    }
}
