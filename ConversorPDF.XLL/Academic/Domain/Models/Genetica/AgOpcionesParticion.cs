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

            string iso = AgDiasVentana.AIso(ventana);
            if (iso == "4,5,6,7")
            {
                bool tiene56M_46T = patrones != null && patrones.Any(p =>
                    !string.IsNullOrWhiteSpace(p) &&
                    (p.IndexOf("6M", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     p.IndexOf("6T", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     string.Equals(p.Trim(), "JUE-DOM", StringComparison.OrdinalIgnoreCase)));

                bool tiene45 = patrones != null && patrones.Any(p => {
                    var d = AgCarriles.Ventana(p);
                    return d.Contains(4) && d.Contains(5);
                });

                bool solo67 = !tiene45 && patrones != null && patrones.Any(p => {
                    var d = AgCarriles.Ventana(p);
                    return d.Contains(6) && d.Contains(7);
                }) && patrones.All(p => string.IsNullOrWhiteSpace(p) || AgCarriles.Ventana(p).All(d => d >= 6));

                if ((tiene45 || solo67) && !tiene56M_46T)
                {
                    bool tiene6T = ranuras.Any(r => r.Dia == 6 &&
                        string.Equals(r.Franja, AgFranjaJornada.TARDE, StringComparison.OrdinalIgnoreCase));
                    string lane45 = tiene6T ? "4,5|6M,6T,7" : "4,5|6,7";
                    var op = Construir(ventana, acc, lane45, ranuras);
                    if (op != null) return op;
                }
                return Construir(ventana, acc, "5,6M|4,6T,7", ranuras);
            }
            // Fase 2441: Catálogo determinista para MIE-DOM (3,4,5,6,7) sin singleton
            if (iso == "3,4,5,6,7")
            {
                bool tiene6T = ranuras.Any(r => r.Dia == 6 &&
                    string.Equals(r.Franja, AgFranjaJornada.TARDE, StringComparison.OrdinalIgnoreCase));
                string laneMieDom = tiene6T ? "5,6M|3,4,6T,7" : "5,6M|3,4,7";
                var op = Construir(ventana, acc, laneMieDom, ranuras);
                if (op != null) return op;
                return Construir(ventana, acc, "5,6|3,4,7", ranuras);
            }

            var origen = DetectarOrigen(patrones);
            if (origen == Origen.Ninguno)
            {
                return null;
            }

            var lane = LaneActual(ventana, origen);
            return Construir(ventana, acc, lane, ranuras);
        }

        /// <summary>
        /// Respaldo de opciones de partición sin bandas singleton para ventanas de 5 días (1..5) cuando
        /// DetectarOrigen devuelve Ninguno debido a patrones mixtos o de 1 día (Fase 2391 / Fase 2434).
        /// Genera opciones binarias {2,3|1,4,5} y {1,2,5|3,4} donde ningún carril es singleton, ordenadas por frecuencia de señales.
        /// </summary>
        internal static IReadOnlyList<AgOpcionParticion> ResolverFallbackSinSingleton(
            string ventana, IReadOnlyList<AgRanura> ranuras, IReadOnlyList<string> patrones)
        {
            string iso = AgDiasVentana.AIso(ventana);
            if (iso == "4,5,6,7")
            {
                bool tiene6T = ranuras != null && ranuras.Any(r => r.Dia == 6 &&
                    string.Equals(r.Franja, AgFranjaJornada.TARDE, StringComparison.OrdinalIgnoreCase));
                string lane45 = tiene6T ? "4,5|6M,6T,7" : "4,5|6,7";
                var op45 = Construir(ventana, "quedarse", lane45, ranuras);
                var opJueDom = Construir(ventana, "quedarse", "5,6M|4,6T,7", ranuras);
                var ops = new List<AgOpcionParticion>();
                if (op45 != null) ops.Add(op45);
                if (opJueDom != null) ops.Add(opJueDom);
                return ops.Count > 0 ? ops : null;
            }

            var dias = AgDiasVentana.Dias(ventana);
            if (dias.Count != 5 || dias[0] != 1 || dias[4] != 5)
                return null;

            int senalesMarVie = 0;
            int senalesLunJue = 0;
            if (patrones != null)
            {
                foreach (var p in patrones.Where(x => !string.IsNullOrWhiteSpace(x)))
                {
                    var dv = AgCarriles.Ventana(p);
                    if (dv.Count == 1)
                    {
                        if (dv[0] == 1 || dv[0] == 4) senalesMarVie++;
                        else if (dv[0] == 5) senalesLunJue++;
                    }
                    else if (dv.Count == 2)
                    {
                        if ((dv[0] == 2 && dv[1] == 3) || (dv[0] == 4 && dv[1] == 5)) senalesMarVie++;
                        else if ((dv[0] == 1 && dv[1] == 2) || (dv[0] == 3 && dv[1] == 4)) senalesLunJue++;
                    }
                }
            }

            var textos = senalesLunJue > senalesMarVie
                ? new[] { "1,2,5|3,4", "2,3|1,4,5" }
                : new[] { "2,3|1,4,5", "1,2,5|3,4" };
            var opciones = new List<AgOpcionParticion>();
            foreach (var texto in textos)
            {
                var op = Construir(ventana, "quedarse", texto, ranuras);
                if (op == null) continue;
                bool sinSingleton = op.Carriles.All(c =>
                    c.Select(r => r.Dia).Distinct().Count() >= 2);
                if (sinSingleton) opciones.Add(op);
            }
            return opciones.Count > 0 ? opciones : null;
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

                // Fase 2434: Patrones de 1 día que son residuales de MarVie o LunJue
                if (dias.Count == 1)
                {
                    if (dias[0] == 1 || dias[0] == 4)
                    {
                        marVie = true;
                        continue;
                    }
                    if (dias[0] == 5)
                    {
                        lunJue = true;
                        continue;
                    }
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
            if (iso == "3,4,5,6,7") return "5,6M|3,4,7"; // Fase 2441: MIE-DOM determinista sin singleton
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
