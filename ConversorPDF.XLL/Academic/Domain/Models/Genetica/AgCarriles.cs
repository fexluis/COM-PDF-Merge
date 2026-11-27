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
        /// Tolera tokens de franja (6M/6T) normalizándolos a día 6 (Fase 2441).
        /// </summary>
        public static IReadOnlyList<int> Ventana(string diasVentana)
        {
            if (string.IsNullOrWhiteSpace(diasVentana)) return new List<int>();
            string norm = diasVentana
                .Replace("6M", "6").Replace("6m", "6")
                .Replace("6T", "6").Replace("6t", "6");
            return AgDiasVentana.Dias(norm);
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
        /// Partición canónica sin singleton acotada exclusivamente a ventanas de 5 días (1..5) (Fase 2392).
        /// Previene la fragmentación en {1,2}|{3,4}|{5} agrupando los días sin singleton ({2,3}|{1,4,5} o {1,2,5}|{3,4}).
        /// Para ventanas de 3 o 7 días retorna null sin alterar su partición canónica.
        /// </summary>
        public static IReadOnlyList<IReadOnlyList<AgRanura>> ParticionSinSingleton(
            IReadOnlyList<AgRanura> ranuras, bool esMarVie)
        {
            if (ranuras == null) return null;

            var dias = ranuras.Select(r => r.Dia).Distinct().OrderBy(d => d).ToList();
            if (dias.Count != 5 || dias[0] != 1 || dias[1] != 2 ||
                dias[2] != 3 || dias[3] != 4 || dias[4] != 5)
                return null;

            var bandasDias = esMarVie
                ? new List<IReadOnlyList<int>> {
                    new List<int> { 2, 3 }, new List<int> { 1, 4, 5 } }
                : new List<IReadOnlyList<int>> {
                    new List<int> { 1, 2, 5 }, new List<int> { 3, 4 } };

            return bandasDias
                .Select(b => (IReadOnlyList<AgRanura>)AgBandasDias.RanurasDe(ranuras, b))
                .ToList();
        }

        /// <summary>
        /// Obtiene la banda sin singleton para un día específico dentro de una ventana de 5 días (1..5) (Fase 2403, Fase 2407).
        /// Valida que la ventana sea exactamente {1, 2, 3, 4, 5}.
        /// MarVie: {1, 4, 5} para 1, 4, 5; {2, 3} para 2, 3.
        /// LunJue: {1, 2, 5} para 1, 2, 5; {3, 4} para 3, 4.
        /// Para ventanas distintas de 5 días, no canónicas o días fuera de 1..5 retorna null.
        /// </summary>
        public static IReadOnlyList<int> BandaSinSingletonParaDia(
            IReadOnlyList<int> ventana, int dia, bool esMarVie)
        {
            if (ventana == null || ventana.Count != 5)
                return null;

            if (ventana[0] != 1 || ventana[1] != 2 || ventana[2] != 3 ||
                ventana[3] != 4 || ventana[4] != 5)
                return null;

            if (dia < 1 || dia > 5) return null;

            if (esMarVie)
            {
                if (dia == 1 || dia == 4 || dia == 5) return new[] { 1, 4, 5 };
                return new[] { 2, 3 };
            }

            if (dia == 1 || dia == 2 || dia == 5) return new[] { 1, 2, 5 };
            return new[] { 3, 4 };
        }

        /// <summary>
        /// Sobrecarga con heurística residual cuando no se especifica origen (Fase 2410, Fase 2415, Fase 2417):
        /// Día 5 -> LunJue ({1, 2, 5}); Día 1 -> MarVie ({1, 4, 5}); Días 2 y 3 -> MarVie ({2, 3}); Día 4 -> MarVie ({1, 4, 5}).
        /// </summary>
        public static IReadOnlyList<int> BandaSinSingletonParaDia(
            IReadOnlyList<int> ventana, int dia)
        {
            if (dia == 5) return BandaSinSingletonParaDia(ventana, 5, false);
            return BandaSinSingletonParaDia(ventana, dia, true);
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
