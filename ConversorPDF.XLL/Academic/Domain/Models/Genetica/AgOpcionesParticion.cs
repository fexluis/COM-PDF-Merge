using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Catálogo determinista de opciones de partición de carriles por origen (Fase 2179 / Fase 2185 / Fase 2186).
    /// Resuelve la partición preferente para una ventana específica y sus ranuras dadas.
    /// Internal para prevenir registro espurio de UDFs en Excel-DNA 1.8.0.
    /// </summary>
    internal static class AgOpcionesParticion
    {
        internal const int MaxOpcionesDecision = 2;
        private static readonly string[] AccionesValidas = { "quedarse", "escalar" };

        internal enum Origen
        {
            Ninguno,
            LunJue,
            MarVie,
            FinDeSemana
        }

        internal static AgOpcionParticion Resolver(
            string ventana, IReadOnlyList<AgRanura> ranuras, IReadOnlyList<string> patrones, string accion = "quedarse")
        {
            string acc = (accion ?? "quedarse").Trim().ToLowerInvariant();
            if (!AccionesValidas.Contains(acc, StringComparer.OrdinalIgnoreCase))
            {
                throw new ArgumentOutOfRangeException(nameof(accion), $"Acción '{accion}' no permitida. Valores válidos: 'quedarse', 'escalar'.");
            }

            if (string.IsNullOrWhiteSpace(ventana) || ranuras == null || ranuras.Count == 0)
            {
                return null;
            }

            var origen = DetectarOrigen(patrones);
            if (origen == Origen.Ninguno)
            {
                return null;
            }

            var lane = LaneActual(ventana, origen);
            return Construir(ventana, acc, lane, ranuras);
        }

        internal static Origen DetectarOrigen(IReadOnlyList<string> patrones)
        {
            if (patrones == null || patrones.Count == 0) return Origen.Ninguno;

            bool marVie = false, lunJue = false, finSemana = false;
            foreach (var p in patrones.Where(x => !string.IsNullOrWhiteSpace(x)))
            {
                string t = p.Trim();
                if (t.IndexOf("6M", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    t.IndexOf("6T", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    finSemana = true;
                    continue;
                }

                if (string.Equals(t, "MAR-VIE", StringComparison.OrdinalIgnoreCase))
                {
                    marVie = true;
                    continue;
                }

                if (string.Equals(t, "LUN-JUE", StringComparison.OrdinalIgnoreCase))
                {
                    lunJue = true;
                    continue;
                }

                var dias = AgCarriles.Ventana(t);
                if (dias.Count == 0) continue;

                if (dias.Any(d => d >= 6))
                {
                    finSemana = true;
                    continue;
                }

                // Carriles canónicos de 2 días de MAR-VIE: {2,3} o {4,5}
                if ((dias.Count == 2 && dias[0] == 2 && dias[1] == 3) ||
                    (dias.Count == 2 && dias[0] == 4 && dias[1] == 5) ||
                    (dias.Count == 4 && dias[0] == 2 && dias[1] == 3 && dias[2] == 4 && dias[3] == 5))
                {
                    marVie = true;
                    continue;
                }

                // Carriles canónicos de 2 días de LUN-JUE: {1,2} o {3,4}
                if ((dias.Count == 2 && dias[0] == 1 && dias[1] == 2) ||
                    (dias.Count == 2 && dias[0] == 3 && dias[1] == 4) ||
                    (dias.Count == 4 && dias[0] == 1 && dias[1] == 2 && dias[2] == 3 && dias[3] == 4))
                {
                    lunJue = true;
                    continue;
                }

                // Patrón terciario (ej. {3,4,5}) u otro no canónico: no clasifica en carril estándar
            }

            if (finSemana) return Origen.FinDeSemana;
            if (marVie && lunJue) return Origen.Ninguno; // mixto: conservador
            if (marVie) return Origen.MarVie;
            if (lunJue) return Origen.LunJue;
            return Origen.Ninguno;
        }

        private static string LaneActual(string v, Origen o)
        {
            string iso = AgDiasVentana.AIso(v);
            if (iso == "2,3,4,5" && o == Origen.MarVie) return "2,3|4,5";
            if (iso == "1,2,3,4" && o == Origen.LunJue) return "1,2|3,4";
            if (iso == "1,2,3,4,5" && o == Origen.MarVie) return "2,3|1,4,5";
            if (iso == "1,2,3,4,5" && o == Origen.LunJue) return "1,2,5|3,4";
            if (iso == "4,5,6,7") return "5,6M|4,6T,7";
            return null;
        }

        private static AgOpcionParticion Construir(
            string ventana,
            string accion,
            string laneStr,
            IReadOnlyList<AgRanura> ranuras)
        {
            if (string.IsNullOrWhiteSpace(laneStr) || ranuras == null || ranuras.Count == 0)
            {
                return null;
            }

            var partes = laneStr.Split('|');
            if (partes.Length < 2)
            {
                return null; // Mínimo 2 carriles
            }

            var carriles = new List<IReadOnlyList<AgRanura>>();
            var ranurasDisponibles = new List<AgRanura>(ranuras);

            foreach (var parte in partes)
            {
                var carril = new List<AgRanura>();
                var tokens = parte.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var tok in tokens)
                {
                    string t = tok.Trim();
                    AgRanura match = null;
                    if (string.Equals(t, "6M", StringComparison.OrdinalIgnoreCase))
                    {
                        match = ranurasDisponibles.FirstOrDefault(r => r.Dia == 6 &&
                            string.Equals(r.Franja, AgFranjaJornada.MANANA, StringComparison.OrdinalIgnoreCase));
                    }
                    else if (string.Equals(t, "6T", StringComparison.OrdinalIgnoreCase))
                    {
                        match = ranurasDisponibles.FirstOrDefault(r => r.Dia == 6 &&
                            string.Equals(r.Franja, AgFranjaJornada.TARDE, StringComparison.OrdinalIgnoreCase));
                    }
                    else if (int.TryParse(t, out int d))
                    {
                        match = ranurasDisponibles.FirstOrDefault(r => r.Dia == d);
                    }

                    if (match == null)
                    {
                        return null; // Token no resoluble en las ranuras provistas
                    }

                    carril.Add(match);
                    ranurasDisponibles.Remove(match);
                }

                if (carril.Count < 2)
                {
                    return null; // CarrilesEstrechos == 0: cada carril debe tener al menos 2 ranuras
                }

                carriles.Add(carril);
            }

            if (ranurasDisponibles.Count != 0)
            {
                return null; // Cobertura exacta de ranuras: no deben sobrar ranuras
            }

            return new AgOpcionParticion(ventana, accion ?? "quedarse", laneStr, carriles);
        }
    }
}
