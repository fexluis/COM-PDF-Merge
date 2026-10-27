using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// V68: Normaliza fechas legacy en ag_sesiones, ag_vinculaciones y ag_gen a formato ISO yyyy-MM-dd con validación de calendario real (Fases 875/879/881/883/885/886/887/890).
    /// Separa la responsabilidad de manipulación de datos de la migración DDL V67.
    /// </summary>
    public sealed class MigrationV68_NormalizarFechasAgSesiones : IMigration, IMigrationTolerante
    {
        public int Version => 68;
        public string Name => "NormalizarFechasAgSesiones";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // Fases 866/868/872/873/875/879/881/883/886: todas las variantes legacy e ISO (con y sin hora/segundos), con validación de calendario real
            var formatos = new[]
            {
                "yyyy-MM-dd", "yyyy-M-d",
                "yyyy-MM-dd HH:mm", "yyyy-MM-dd H:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd H:mm:ss",
                "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss.fff",
                "dd/MM/yyyy", "d/M/yyyy", "dd/M/yyyy", "d/MM/yyyy",
                "dd/MM/yyyy HH:mm", "dd/MM/yyyy H:mm", "dd/MM/yyyy HH:mm:ss", "dd-MM-yyyy", "yyyy/MM/dd"
            };

            // Fase 890: Métrica de truncamiento de marcas de tiempo a solo fecha yyyy-MM-dd
            int truncadas = 0;

            // 1. Normalización en ag_sesiones
            int sesExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_sesiones';", transaction: transaction);
            if (sesExists > 0)
            {
                // Fase 881: date() en SQLite renormaliza fechas imposibles (ej 2026-02-31 -> 2026-03-03);
                // comparando date(f) <> f detectamos fechas ISO inválidas que requieren validación.
                var pendientes = connection.Query(
                    "SELECT id AS Id, fechasesion AS Fecha FROM ag_sesiones " +
                    "WHERE TRIM(IFNULL(fechasesion,'')) <> '' " +
                    "  AND (date(fechasesion) IS NULL OR date(fechasesion) <> fechasesion);",
                    transaction: transaction).Cast<IDictionary<string, object>>().ToList();

                foreach (var fila in pendientes)
                {
                    string raw = Convert.ToString(fila["Fecha"]);
                    long id = Convert.ToInt64(fila["Id"]);
                    string iso = Normalizar(raw, formatos);
                    if (iso != null)
                    {
                        if (raw.Trim().Length > 10) truncadas++;
                        connection.Execute("UPDATE ag_sesiones SET fechasesion = @F WHERE id = @Id;",
                            new { F = iso, Id = id }, transaction);
                    }
                }

                // Fases 881, 883 y 886: Diagnóstico de residuos no normalizables o fechas imposibles tras el barrido
                int residuales = connection.ExecuteScalar<int>(@"
                    SELECT COUNT(*) FROM ag_sesiones
                     WHERE TRIM(IFNULL(fechasesion,'')) <> ''
                       AND (length(fechasesion) <> 10
                            OR date(fechasesion) IS NULL
                            OR date(fechasesion) <> fechasesion);",
                    transaction: transaction);

                if (residuales > 0)
                {
                    Core.Helpers.LoggerHelper.Warning(
                        $"[V68] {residuales} sesión(es) con fecha no normalizable o fecha imposible; DiaSemana quedará NULL y la columna K se resolverá mediante fallback de exportación.");
                }
            }

            // 2. Fase 879, 881 y 887: Normalización en ag_vinculaciones (fecha_inicio y fecha_fin)
            int vincExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_vinculaciones';", transaction: transaction);
            int resVinc = 0;
            if (vincExists > 0)
            {
                var pendientesVinc = connection.Query(
                    "SELECT id AS Id, fecha_inicio AS Ini, fecha_fin AS Fin FROM ag_vinculaciones " +
                    "WHERE (TRIM(IFNULL(fecha_inicio,'')) NOT IN ('','--') " +
                    "       AND (date(fecha_inicio) IS NULL OR date(fecha_inicio) <> fecha_inicio)) " +
                    "   OR (TRIM(IFNULL(fecha_fin,''))    NOT IN ('','--') " +
                    "       AND (date(fecha_fin)    IS NULL OR date(fecha_fin)    <> fecha_fin));",
                    transaction: transaction).Cast<IDictionary<string, object>>().ToList();

                foreach (var fila in pendientesVinc)
                {
                    long id = Convert.ToInt64(fila["Id"]);
                    string ini = Convert.ToString(fila["Ini"]);
                    string fin = Convert.ToString(fila["Fin"]);

                    string iniIso = Normalizar(ini, formatos);
                    string finIso = Normalizar(fin, formatos);
                    if (iniIso == null && finIso == null) continue;

                    if (iniIso != null && ini.Trim().Length > 10) truncadas++;
                    if (finIso != null && fin.Trim().Length > 10) truncadas++;

                    connection.Execute(
                        "UPDATE ag_vinculaciones SET fecha_inicio = @I, fecha_fin = @F WHERE id = @Id;",
                        new { I = iniIso ?? ini, F = finIso ?? fin, Id = id }, transaction);
                }

                // Fase 887: Conteo de residuales en ag_vinculaciones
                resVinc = connection.ExecuteScalar<int>(@"
                    SELECT COUNT(*) FROM ag_vinculaciones
                     WHERE (TRIM(IFNULL(fecha_inicio,'')) NOT IN ('','--')
                            AND (length(fecha_inicio) <> 10 OR date(fecha_inicio) IS NULL OR date(fecha_inicio) <> fecha_inicio))
                        OR (TRIM(IFNULL(fecha_fin,'')) NOT IN ('','--')
                            AND (length(fecha_fin) <> 10 OR date(fecha_fin) IS NULL OR date(fecha_fin) <> fecha_fin));",
                    transaction: transaction);
            }

            // 3. Fase 885 y 887: Normalización en ag_gen (fecha_inicio_estimada y fecha_fin_estimada)
            int genExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_gen';", transaction: transaction);
            int resGen = 0;
            if (genExists > 0)
            {
                var pendientesGen = connection.Query(
                    "SELECT id AS Id, fecha_inicio_estimada AS Ini, fecha_fin_estimada AS Fin FROM ag_gen " +
                    "WHERE (TRIM(IFNULL(fecha_inicio_estimada,'')) <> '' " +
                    "       AND (date(fecha_inicio_estimada) IS NULL OR date(fecha_inicio_estimada) <> fecha_inicio_estimada)) " +
                    "   OR (TRIM(IFNULL(fecha_fin_estimada,'')) <> '' " +
                    "       AND (date(fecha_fin_estimada) IS NULL OR date(fecha_fin_estimada) <> fecha_fin_estimada));",
                    transaction: transaction).Cast<IDictionary<string, object>>().ToList();

                foreach (var fila in pendientesGen)
                {
                    long id = Convert.ToInt64(fila["Id"]);
                    string ini = Convert.ToString(fila["Ini"]);
                    string fin = Convert.ToString(fila["Fin"]);

                    string iniIso = Normalizar(ini, formatos);
                    string finIso = Normalizar(fin, formatos);
                    if (iniIso == null && finIso == null) continue;

                    if (iniIso != null && ini.Trim().Length > 10) truncadas++;
                    if (finIso != null && fin.Trim().Length > 10) truncadas++;

                    connection.Execute(
                        "UPDATE ag_gen SET fecha_inicio_estimada = @I, fecha_fin_estimada = @F WHERE id = @Id;",
                        new { I = iniIso ?? ini, F = finIso ?? fin, Id = id }, transaction);
                }

                // Fase 887: Conteo de residuales en ag_gen
                resGen = connection.ExecuteScalar<int>(@"
                    SELECT COUNT(*) FROM ag_gen
                     WHERE (TRIM(IFNULL(fecha_inicio_estimada,'')) <> ''
                            AND (length(fecha_inicio_estimada) <> 10 OR date(fecha_inicio_estimada) IS NULL
                                 OR date(fecha_inicio_estimada) <> fecha_inicio_estimada))
                        OR (TRIM(IFNULL(fecha_fin_estimada,'')) <> ''
                            AND (length(fecha_fin_estimada) <> 10 OR date(fecha_fin_estimada) IS NULL
                                 OR date(fecha_fin_estimada) <> fecha_fin_estimada));",
                    transaction: transaction);
            }

            // Fase 890: Registro informativo de métrica de truncamiento si se descartaron componentes horarias
            if (truncadas > 0)
            {
                Core.Helpers.LoggerHelper.Info(
                    $"[V68] {truncadas} fecha(s) con hora truncadas a yyyy-MM-dd.");
            }

            if (resVinc > 0 || resGen > 0)
            {
                Core.Helpers.LoggerHelper.Warning(
                    $"[V68] Residuales: ag_vinculaciones={resVinc}, ag_gen={resGen}. Revise formato/calendario antes de promover.");
            }
        }

        private static string Normalizar(string raw, string[] formatos)
        {
            if (string.IsNullOrWhiteSpace(raw) || raw == "--") return null;
            return DateTime.TryParseExact(raw.Trim(), formatos, CultureInfo.InvariantCulture,
                       DateTimeStyles.None, out var dt) ? dt.ToString("yyyy-MM-dd") : null;
        }
    }
}
