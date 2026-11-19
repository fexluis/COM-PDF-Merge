using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Clasificador canónico y unificado del origen de vacantes para el AG (Fase 2092, Fase 2097, Fase 2098).
    /// Centraliza la distinción entre vacantes estructurales docentes, no docentes, liberadas por el AG e indeterminadas.
    /// </summary>
    public static class AgClasificadorVacante
    {
        /// <summary>
        /// Genera la clave de cohorte canónica discriminando por aula/gen si faltan semestre o grupo (Fase 2097).
        /// </summary>
        public static string ClaveCohorte(AgGen g)
        {
            if (g == null) return string.Empty;
            return string.IsNullOrWhiteSpace(g.Semestre) || string.IsNullOrWhiteSpace(g.Grupo)
                ? $"{g.IdAsignatura}|{g.IdSede}|aula:{g.IdAula}|gen:{g.Id}"
                : $"{g.IdAsignatura}|{g.IdSede}|{g.Semestre}|{g.Grupo}";
        }

        /// <summary>
        /// Clasifica el motivo textual o código de vacante en su origen canónico (Fase 2092, Fase 2098).
        /// </summary>
        /// <param name="motivo">Código o descripción del motivo de vacante.</param>
        /// <returns>Categoría canónica del origen de la vacante.</returns>
        public static AgOrigenVacante Clasificar(string motivo)
        {
            var cod = AgMotivoVacanteExtensions.Desde(motivo);
            if (cod == AgMotivoVacante.Ninguno)
            {
                return AgOrigenVacante.Indeterminado;
            }

            if (cod == AgMotivoVacante.SIN_PLAN_OLEADA || cod == AgMotivoVacante.SIN_BLOQUE_NIVEL)
            {
                return AgOrigenVacante.NoDocente;
            }

            if (cod == AgMotivoVacante.SIN_CANDIDATOS || cod == AgMotivoVacante.NO_CUMPLE_REQUISITOS
                || cod == AgMotivoVacante.POOL_VACIO)
            {
                return AgOrigenVacante.EstructuralDocente;
            }

            return AgOrigenVacante.LiberadaAg;
        }
    }
}
