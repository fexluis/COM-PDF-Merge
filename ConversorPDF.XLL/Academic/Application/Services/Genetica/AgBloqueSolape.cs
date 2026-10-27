using System;
using System.Globalization;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Utilitario para cálculo matemático de solapamiento temporal entre bloques horarios (Fase 950 / Fase 957).
    /// Evalúa día de la semana e intersección de intervalos horarios (i1 &lt; f2 &amp;&amp; i2 &lt; f1).
    /// </summary>
    public static class AgBloqueSolape
    {
        public static bool SeSolapan(AgBloqueHorario a, AgBloqueHorario b)
        {
            if (a == null || b == null || a.DiaSemanaNumero != b.DiaSemanaNumero)
            {
                return false;
            }

            if (!TimeSpan.TryParse(a.HoraInicio, CultureInfo.InvariantCulture, out var i1)
             || !TimeSpan.TryParse(a.HoraFin, CultureInfo.InvariantCulture, out var f1)
             || !TimeSpan.TryParse(b.HoraInicio, CultureInfo.InvariantCulture, out var i2)
             || !TimeSpan.TryParse(b.HoraFin, CultureInfo.InvariantCulture, out var f2))
            {
                return false;
            }

            return i1 < f2 && i2 < f1;
        }
    }
}
