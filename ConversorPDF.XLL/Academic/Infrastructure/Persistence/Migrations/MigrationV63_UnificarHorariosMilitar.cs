using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.Services;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V63: Unificación idempotente de horarios al formato militar "HH:mm A HH:mm"
    /// en sesiones.horariosesion, ag_sesiones.horariosesion y vinculaciones.dia_horas_semanales.
    /// Preserva intactas las fechas, horas y duraciones.
    /// </summary>
    public class MigrationV63_UnificarHorariosMilitar : IMigration, IMigrationTolerante
    {
        public int Version => 63;
        public string Name => "UnificarHorariosMilitar";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            LoggerHelper.Info("[V63] Iniciando unificación militar de horarios...");

            // 1. Unificar sesiones.horariosesion
            try
            {
                var sesiones = connection.Query<(int Id, string Horario)>(
                    "SELECT id AS Id, horariosesion AS Horario FROM sesiones WHERE horariosesion IS NOT NULL AND horariosesion <> '';",
                    transaction: transaction).ToList();

                int actualizadasSesiones = 0;
                foreach (var s in sesiones)
                {
                    string normalizado = HorarioMilitarizador.NormalizarHorario(s.Horario);
                    if (!string.Equals(normalizado, s.Horario, StringComparison.Ordinal))
                    {
                        connection.Execute(
                            "UPDATE sesiones SET horariosesion = @Horario WHERE id = @Id;",
                            new { Horario = normalizado, Id = s.Id },
                            transaction: transaction);
                        actualizadasSesiones++;
                    }
                }
                LoggerHelper.Info($"[V63] 'sesiones.horariosesion' normalizadas: {actualizadasSesiones} de {sesiones.Count}.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V63] Error o tabla ausente en sesiones: {ex.Message}");
            }

            // 2. Unificar ag_sesiones.horariosesion
            try
            {
                int existeAgSesiones = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_sesiones';",
                    transaction: transaction);

                if (existeAgSesiones > 0)
                {
                    var agSesiones = connection.Query<(int Id, string Horario)>(
                        "SELECT id AS Id, horariosesion AS Horario FROM ag_sesiones WHERE horariosesion IS NOT NULL AND horariosesion <> '';",
                        transaction: transaction).ToList();

                    int actualizadasAg = 0;
                    foreach (var s in agSesiones)
                    {
                        string normalizado = HorarioMilitarizador.NormalizarHorario(s.Horario);
                        if (!string.Equals(normalizado, s.Horario, StringComparison.Ordinal))
                        {
                            connection.Execute(
                                "UPDATE ag_sesiones SET horariosesion = @Horario WHERE id = @Id;",
                                new { Horario = normalizado, Id = s.Id },
                                transaction: transaction);
                            actualizadasAg++;
                        }
                    }
                    LoggerHelper.Info($"[V63] 'ag_sesiones.horariosesion' normalizadas: {actualizadasAg} de {agSesiones.Count}.");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V63] Error o tabla ausente en ag_sesiones: {ex.Message}");
            }

            // 3. Unificar vinculaciones.dia_horas_semanales
            try
            {
                var vinculaciones = connection.Query<(int Id, string DiaHoras)>(
                    "SELECT id AS Id, dia_horas_semanales AS DiaHoras FROM vinculaciones WHERE dia_horas_semanales IS NOT NULL AND dia_horas_semanales <> '';",
                    transaction: transaction).ToList();

                int actualizadasVinc = 0;
                foreach (var v in vinculaciones)
                {
                    // Buscar horario modal de las sesiones vinculadas
                    string horarioModal = connection.QueryFirstOrDefault<string>(@"
SELECT horariosesion FROM sesiones 
WHERE idvinculacion = @IdVinculacion AND horariosesion IS NOT NULL AND horariosesion <> ''
GROUP BY horariosesion 
ORDER BY COUNT(*) DESC 
LIMIT 1;", new { IdVinculacion = v.Id }, transaction: transaction);

                    string horarioModalNorm = HorarioMilitarizador.NormalizarHorario(horarioModal);

                    var partesEntrada = v.DiaHoras.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    var entradasProcesadas = new List<string>();

                    foreach (var pe in partesEntrada)
                    {
                        var segs = pe.Trim().Split('|');
                        if (segs.Length == 3)
                        {
                            string dia = segs[0].Trim();
                            string horas = segs[1].Trim();
                            string horario = segs[2].Trim();

                            string horarioFinal = !string.IsNullOrEmpty(horario)
                                ? HorarioMilitarizador.NormalizarHorario(horario)
                                : horarioModalNorm;

                            entradasProcesadas.Add($"{dia}|{horas}|{horarioFinal}");
                        }
                        else if (segs.Length == 2)
                        {
                            string dia = segs[0].Trim();
                            string horas = segs[1].Trim();
                            entradasProcesadas.Add($"{dia}|{horas}|{horarioModalNorm}");
                        }
                        else
                        {
                            entradasProcesadas.Add(pe.Trim());
                        }
                    }

                    string nuevoDiaHoras = string.Join(", ", entradasProcesadas);
                    if (!string.Equals(nuevoDiaHoras, v.DiaHoras, StringComparison.Ordinal))
                    {
                        connection.Execute(
                            "UPDATE vinculaciones SET dia_horas_semanales = @DiaHoras WHERE id = @Id;",
                            new { DiaHoras = nuevoDiaHoras, Id = v.Id },
                            transaction: transaction);
                        actualizadasVinc++;
                    }
                }
                LoggerHelper.Info($"[V63] 'vinculaciones.dia_horas_semanales' normalizadas: {actualizadasVinc} de {vinculaciones.Count}.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[V63] Error o tabla ausente en vinculaciones: {ex.Message}");
            }

            LoggerHelper.Info("[V63] Unificación militar completada.");
        }
    }
}
