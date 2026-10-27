using System.Globalization;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Generador determinista de clave primaria de 63 bits para la entidad <see cref="TerceroAsignatura"/>.
    /// Implementa el algoritmo FNV-1a de 64 bits enmascarado a 63 bits (entero positivo con signo para SQLite)
    /// garantizando unicidad y determinismo para la tupla (IdTercero, IdPrograma, IdAsignatura, Semestre).
    /// </summary>
    public static class TerceroAsignaturaKey
    {
        private const ulong FnvOffsetBasis = 14695981039346656037UL;
        private const ulong FnvPrime = 1099511628211UL;
        private const ulong Mask63Bits = 0x7FFFFFFFFFFFFFFFUL;

        /// <summary>
        /// Construye la clave lógica canónica en formato texto (pre-hash) para auditoría y detección de colisiones.
        /// </summary>
        /// <param name="idTercero">Identificador del tercero/docente.</param>
        /// <param name="idPrograma">Identificador del programa académico.</param>
        /// <param name="idAsignatura">Identificador de la asignatura.</param>
        /// <param name="semestre">Semestre lectivo o cadena vacía/null si es transversal.</param>
        /// <returns>Cadena concatenada canónica: "{idTercero}|{idPrograma}|{idAsignatura}|{semestreSanitizado}".</returns>
        public static string ConstruirClave(int idTercero, int idPrograma, int idAsignatura, string semestre)
        {
            var ci = CultureInfo.InvariantCulture;
            string cleanSemestre = (semestre ?? string.Empty).Trim().ToUpperInvariant();
            return string.Concat(
                idTercero.ToString(ci), "|",
                idPrograma.ToString(ci), "|",
                idAsignatura.ToString(ci), "|",
                cleanSemestre);
        }

        /// <summary>
        /// Calcula la clave primaria determinista de 63 bits a partir de los componentes de la tupla.
        /// </summary>
        /// <param name="idTercero">Identificador del tercero/docente.</param>
        /// <param name="idPrograma">Identificador del programa académico.</param>
        /// <param name="idAsignatura">Identificador de la asignatura.</param>
        /// <param name="semestre">Semestre lectivo o cadena vacía/null si es transversal.</param>
        /// <returns>Entero de 64 bits positivo (<see cref="long"/>) para ser usado como PK física sin AUTOINCREMENT.</returns>
        public static long Calcular(int idTercero, int idPrograma, int idAsignatura, string semestre)
        {
            string keyString = ConstruirClave(idTercero, idPrograma, idAsignatura, semestre);

            unchecked
            {
                ulong hash = FnvOffsetBasis;
                foreach (char c in keyString)
                {
                    hash ^= c;
                    hash *= FnvPrime;
                }
                return (long)(hash & Mask63Bits);
            }
        }
    }
}
