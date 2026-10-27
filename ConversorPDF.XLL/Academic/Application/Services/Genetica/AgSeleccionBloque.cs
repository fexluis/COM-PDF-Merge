using System;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Helper estático para la selección determinista y unificada de bloques horarios (Fase 1135).
    /// Resuelve de forma consistente día, duración y franja canónica de la jornada del aula en
    /// fábrica, cruce, mutación y coherencia.
    /// </summary>
    internal static class AgSeleccionBloque
    {
        public static AgBloqueHorario ParaDia(AgAula aula, int dia, double unidad, AgContexto ctx, string franjaPreferida = null)
        {
            string franja = !string.IsNullOrWhiteSpace(franjaPreferida)
                ? franjaPreferida
                : (aula != null && !string.IsNullOrWhiteSpace(aula.Jornada)
                    ? AgFranjaJornada.Resolve(aula.Jornada, (int)unidad, dia)?.Franja : null);

            if (ctx != null && ctx.TryGetBloque(dia, unidad, franja, out var b)) return b;

            // Fallback: primera franja válida de la ranura del día (Fase 1160)
            var franjasDia = AgRanuras.FranjasDe(dia, aula?.Jornada, unidad);
            foreach (var f in franjasDia)
            {
                if (ctx != null && ctx.TryGetBloque(dia, unidad, f, out var bf)) return bf;
            }

            return null;
        }
    }
}
