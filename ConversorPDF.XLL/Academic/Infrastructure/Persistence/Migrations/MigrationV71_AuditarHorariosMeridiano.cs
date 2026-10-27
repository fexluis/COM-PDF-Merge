using System;
using System.Data;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V71: Audita y reporta registros de horarios afectados por la normalización previa con 'hh' (Fase 1000).
    /// Migración no destructiva de solo diagnóstico y advertencia para guiar la revisión operativa.
    /// </summary>
    public sealed class MigrationV71_AuditarHorariosMeridiano : IMigration, IMigrationTolerante
    {
        public int Version => 71;
        public string Name => "AuditarHorariosMeridiano";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            try
            {
                int existeAgSesiones = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_sesiones';",
                    transaction: transaction);
                int existeSesiones = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='sesiones';",
                    transaction: transaction);

                // Fase 1005: criterio de auditoría preciso comparando contra ag_bloque_horario
                string sqlSospechosos = @"
SELECT COUNT(*) FROM {0} s
WHERE COALESCE(s.horariosesion,'') <> ''
  AND EXISTS (SELECT 1 FROM ag_bloque_horario LIMIT 1)
  AND NOT EXISTS (
        SELECT 1 FROM ag_bloque_horario b
        WHERE b.hora_inicio || ' A ' || b.hora_fin = s.horariosesion);";

                int vincAg = existeAgSesiones > 0 ? connection.ExecuteScalar<int>(string.Format(sqlSospechosos, "ag_sesiones"), transaction: transaction) : 0;
                int vincProd = existeSesiones > 0 ? connection.ExecuteScalar<int>(string.Format(sqlSospechosos, "sesiones"), transaction: transaction) : 0;

                LoggerHelper.Warning(
                    $"[V71] Horarios afectados por V63 (hh→HH) requieren revisión manual: ag_sesiones={vincAg}, sesiones={vincProd}. " +
                    "Regenere corridas AG (borrar ag_corrida/ag_vinculaciones/ag_sesiones) y revalide producción contra dia_horas_semanales.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V71] Error en diagnóstico de auditoría de horarios: {ex.Message}");
            }
        }
    }
}
