using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Core.Models.Academic;

namespace ConversorPDF.XLL.Core.Services.Academic
{
    public class GeneradorSesionesRecurrentes
    {
        /// <summary>
        /// Genera sesiones semanales según configuración de días consecutivos
        /// </summary>
        public List<Sesion> Generar(
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
                return sesiones;

            // Validar días consecutivos (ej: Viernes-Sábado)
            bool sonConsecutivos = SonDiasConsecutivos(diasConfig);
            if (!sonConsecutivos && diasConfig.Count > 1)
                throw new ArgumentException("Los días configurados deben ser consecutivos para unir la sesión en una sola fila.");
            
            // Calcular horas por sesión (suma de todas las configuraciones)
            double horasPorSesion = diasConfig.Sum(d => (double)d.Horas);
            double horasAcumuladas = 0;
            
            // Encontrar primera ocurrencia del primer día
            var primerDia = diasConfig.OrderBy(d => d.DiaSemana).First();
            DateTime fechaBase = EncontrarPrimeraFecha(fechaInicioGlobal, primerDia.DiaSemana);
            
            // Generar sesiones semanales recurrentes
            while (horasAcumuladas < horasTotalesRequeridas)
            {
                DateTime fechaIni = fechaBase.AddDays((sesiones.Count) * 7);
                DateTime fechaFin = fechaIni.AddDays(diasConfig.Count - 1);
                
                // Validar que no exceda el rango global
                if (fechaFin > fechaFinGlobal)
                    break;
                
                // Si la última sesión excede el total, ajustar las horas
                double horasEstaSesion = horasPorSesion;
                if (horasAcumuladas + horasEstaSesion > horasTotalesRequeridas)
                {
                    horasEstaSesion = horasTotalesRequeridas - horasAcumuladas;
                }

                // Calcular duración: (días de diferencia) + 0.5
                double duracion = (fechaFin - fechaIni).Days + 0.5;
                
                sesiones.Add(new Sesion
                {
                    Id = 0,
                    IdVinculacion = idVinculacion,
                    IdTercero = idTercero,
                    fechasesionInicial = fechaIni,
                    fechasesionFinal = fechaFin,
                    Duracion = duracion,
                    Horas = horasEstaSesion,
                    Viatica = viaticaDefault
                });

                horasAcumuladas += horasEstaSesion;
                
                if (horasAcumuladas >= horasTotalesRequeridas)
                    break;
            }
            
            return sesiones;
        }
        
        private bool SonDiasConsecutivos(List<ConfiguracionDia> dias)
        {
            if (dias.Count < 2) return true;
            var ordenados = dias.OrderBy(d => d.DiaSemana).ToList();
            
            for (int i = 1; i < ordenados.Count; i++)
            {
                int diff = ((int)ordenados[i].DiaSemana - (int)ordenados[i-1].DiaSemana + 7) % 7;
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
