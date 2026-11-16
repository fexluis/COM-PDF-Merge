using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Prioridad canónica de asignación por horas efectivas, orden curricular e identificador.</summary>
    internal static class AgOrdenAsignaturas
    {
        /// <summary>Ordena sin modificar las demandas ni sus horas originales.</summary>
        internal static List<(int IdAsignatura, double Horas, int OrdenCurricular)> PorHoras(
            IEnumerable<(int IdAsignatura, double Horas, int OrdenCurricular)> asignaturas)
        {
            return asignaturas.OrderByDescending(a => a.Horas)
                .ThenBy(a => Math.Max(0, a.OrdenCurricular))
                .ThenBy(a => a.IdAsignatura).ToList();
        }
    }
}
