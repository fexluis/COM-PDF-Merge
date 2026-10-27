using System;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Catálogo y clasificador del nivel académico a partir de la modalidad o denominación del programa (Fase 936 / D17).
    /// </summary>
    public static class AgNivelAcademicoCatalogo
    {
        public static AgNivelAcademico Desde(string modalidad)
        {
            if (string.IsNullOrWhiteSpace(modalidad))
            {
                return AgNivelAcademico.Pregrado;
            }

            string upper = modalidad.Trim().ToUpperInvariant();
            if (upper.Contains("POSGRADO") ||
                upper.Contains("ESPECIALIZACION") ||
                upper.Contains("ESPECIALIZACIÓN") ||
                upper.Contains("MAESTRIA") ||
                upper.Contains("MAESTRÍA") ||
                upper.Contains("DOCTORADO"))
            {
                return AgNivelAcademico.Posgrado;
            }

            return AgNivelAcademico.Pregrado;
        }
    }
}
