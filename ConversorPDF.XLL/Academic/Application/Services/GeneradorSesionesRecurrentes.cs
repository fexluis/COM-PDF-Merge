using System;
using System.Collections.Generic;
using System.Linq;
using CSharpFunctionalExtensions;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    public class GeneradorSesionesRecurrentes
    {
        /// <summary>
        /// Genera sesiones semanales según configuración de días consecutivos.
        /// </summary>
        /// <returns>
        /// <see cref="Result{T}"/> con la lista de sesiones generadas, o un error
        /// si los días configurados no son consecutivos.
        /// </returns>
        public Result<IReadOnlyList<Sesion>> Generar(
            int idTercero,
            int idVinculacion,
            DateTime fechaInicioGlobal,
            DateTime fechaFinGlobal,
            List<ConfiguracionDia> diasConfig,
            double horasTotalesRequeridas,
            bool viaticaDefault)
        {
            var sesiones = new List<Sesion>();

            if (diasConfig == null || !diasConfig.Any())
                return Result.Success<IReadOnlyList<Sesion>>(sesiones);

            // Validar días consecutivos (ej: Viernes-Sábado)
            bool sonConsecutivos = SonDiasConsecutivos(diasConfig);
            if (!sonConsecutivos && diasConfig.Count > 1)
                return Result.Failure<IReadOnlyList<Sesion>>(
                    "Los días configurados deben ser consecutivos para unir la sesión en una sola fila.");

            double horasAcumuladas = 0;

            // Encontrar primera ocurrencia del primer día
            var primerDia = diasConfig.OrderBy(d => d.DiaSemana).First();
            DateTime fechaBase = EncontrarPrimeraFecha(fechaInicioGlobal, primerDia.DiaSemana);

            // Generar sesiones semanales recurrentes
            while (horasAcumuladas < horasTotalesRequeridas)
            {
                DateTime fechaBase2 = fechaBase.AddDays((sesiones.Count / diasConfig.Count) * 7);

                foreach (var diaConf in diasConfig.OrderBy(d => d.DiaSemana))
                {
                    if (horasAcumuladas >= horasTotalesRequeridas)
                        break;

                    int diasDesdeBase = ((int)diaConf.DiaSemana - (int)diasConfig.OrderBy(d => d.DiaSemana).First().DiaSemana + 7) % 7;
                    DateTime fechaDia = fechaBase2.AddDays(diasDesdeBase);

                    if (fechaDia > fechaFinGlobal)
                    {
                        horasAcumuladas = horasTotalesRequeridas; // Forzar salida del while
                        break;
                    }

                    double horasEstaSesion = (double)diaConf.Horas;
                    double horasFaltantes = horasTotalesRequeridas - horasAcumuladas;
                    if (horasEstaSesion > horasFaltantes)
                    {
                        horasEstaSesion = horasFaltantes;
                    }

                    string nombreDia = diaConf.Nombre?.ToUpper()
                        ?? fechaDia.ToString("dddd", System.Globalization.CultureInfo.GetCultureInfo("es-ES")).ToUpper();

                    sesiones.Add(new Sesion
                    {
                        Id = 0,
                        IdVinculacion = idVinculacion,
                        IdTercero = idTercero,
                        FechaSesion = fechaDia,
                        HorarioSesion = !string.IsNullOrEmpty(diaConf.Horario) ? diaConf.Horario : $"{nombreDia}|{(int)diaConf.Horas}",
                        Horas = horasEstaSesion,
                        Viatica = viaticaDefault
                    });

                    horasAcumuladas += horasEstaSesion;
                }

                if (horasAcumuladas >= horasTotalesRequeridas)
                    break;
            }

            return Result.Success<IReadOnlyList<Sesion>>(sesiones);
        }

        private bool SonDiasConsecutivos(List<ConfiguracionDia> dias)
        {
            if (dias.Count < 2) return true;
            var ordenados = dias.OrderBy(d => d.DiaSemana).ToList();

            for (int i = 1; i < ordenados.Count; i++)
            {
                int diff = ((int)ordenados[i].DiaSemana - (int)ordenados[i - 1].DiaSemana + 7) % 7;
                if (diff != 1) return false;
            }
            return true;
        }

        private DateTime EncontrarPrimeraFecha(DateTime inicio, DayOfWeek diaObjetivo)
        {
            int diasDiff = ((int)diaObjetivo - (int)inicio.DayOfWeek + 7) % 7;
            // Si el día de inicio es el mismo que el objetivo, no sumamos 7 a menos que el requerimiento lo exija.
            // Según el ejemplo: 16/04/2026 es jueves, empieza viernes 17/04/2026. (diasDiff = 1)
            return inicio.AddDays(diasDiff);
        }
    }
}
