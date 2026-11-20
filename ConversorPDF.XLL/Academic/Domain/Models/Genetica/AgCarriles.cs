using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Utilidades puras para partición, cálculo y rotación modular de carriles disjuntos (Fase 1055, Fase 1154).
    /// Internal para prevenir registro espurio de UDFs en Excel-DNA 1.8.0.
    /// </summary>
    internal static class AgCarriles
    {
        public const int RanurasPorCarril = 2;
        public const int DiasPorCarril = 2;

        /// <summary>
        /// Días ISO de la ventana semanal; vacío si no hay ventana declarada.
        /// </summary>
        public static IReadOnlyList<int> Ventana(string diasVentana)
        {
            return AgDiasVentana.Dias(diasVentana);
        }

        /// <summary>
        /// Cantidad de carriles disjuntos de 2 ranuras que caben en la lista de ranuras. Mínimo 1 (Fase 1154).
        /// </summary>
        public static int Cantidad(IReadOnlyList<AgRanura> ranuras)
        {
            if (ranuras == null || ranuras.Count == 0) return 1;
            return Math.Max(1, (int)Math.Ceiling(ranuras.Count / (double)RanurasPorCarril));
        }

        /// <summary>
        /// Cantidad de carriles disjuntos de 2 días que caben en la ventana. Mínimo 1.
        /// </summary>
        public static int Cantidad(IReadOnlyList<int> ventana)
        {
            if (ventana == null || ventana.Count == 0) return 1;
            return Math.Max(1, (int)Math.Ceiling(ventana.Count / (double)DiasPorCarril));
        }

        /// <summary>
        /// Rotación modular de carril: nunca negativo, acotado a [0, carriles - 1].
        /// </summary>
        public static int Indice(int idxAsignatura, int carriles)
        {
            return Math.Abs(idxAsignatura) % Math.Max(1, carriles);
        }

        /// <summary>
        /// Obtiene el subconjunto disjunto de 2 ranuras correspondiente al carril rotacional de la asignatura (Fase 1154).
        /// </summary>
        public static IReadOnlyList<AgRanura> DelCarril(IReadOnlyList<AgRanura> ranuras, int idxAsignatura)
        {
            if (ranuras == null || ranuras.Count == 0)
            {
                return new List<AgRanura>();
            }

            int carriles = Cantidad(ranuras);
            int idx = Indice(idxAsignatura, carriles);
            return ranuras.Skip(idx * RanurasPorCarril).Take(RanurasPorCarril).ToList();
        }

        /// <summary>
        /// Obtiene el subconjunto disjunto de 2 días correspondiente al carril rotacional de la asignatura.
        /// </summary>
        public static IReadOnlyList<int> DelCarril(IReadOnlyList<int> ventana, int idxAsignatura)
        {
            if (ventana == null || ventana.Count == 0)
            {
                return new List<int>();
            }

            int carriles = Cantidad(ventana);
            int idx = Indice(idxAsignatura, carriles);
            return ventana.Skip(idx * DiasPorCarril).Take(DiasPorCarril).ToList();
        }

        /// <summary>
        /// Genera la clave canónica de carril por aula y patrón semanal (Fase 2140 / Fase 2156).
        /// </summary>
        internal static string ClaveCarril(AgGen g) =>
            $"{g.IdAula}|{(string.IsNullOrWhiteSpace(g.DiasPatron) ? "DEF" : g.DiasPatron)}";

        /// <summary>
        /// Días residuales de ventana impar (por ejemplo, en ventana de 5 días, el 5º día que queda sin rotar).
        /// </summary>
        [System.Obsolete("F1176: el modelo de ranuras no descarta días; usar AgCarriles.DelCarril(ranuras, idx).")]
        public static IReadOnlyList<int> Residual(IReadOnlyList<int> ventana)
        {
            if (ventana == null || ventana.Count == 0)
            {
                return new List<int>();
            }

            int carriles = Cantidad(ventana);
            return ventana.Skip(carriles * DiasPorCarril).ToList();
        }
    }
}
