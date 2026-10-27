using System;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Servicio de resolución de duración y unidades de sesión académica según nivel y aula (Fase 936 / D4 / D4').
    /// </summary>
    public static class AgUnidadesSemanales
    {
        public static double HorasSesionAula(AgAula aula, AgNivelAcademico nivel)
        {
            if (nivel == AgNivelAcademico.Posgrado && aula != null && aula.HorasDia == 8)
            {
                return 8.0;
            }

            return 4.0;
        }

        public static double HorasSesionAula(int horasDia, AgNivelAcademico nivel)
        {
            if (nivel == AgNivelAcademico.Posgrado && horasDia == 8)
            {
                return 8.0;
            }

            return 4.0;
        }
    }
}
