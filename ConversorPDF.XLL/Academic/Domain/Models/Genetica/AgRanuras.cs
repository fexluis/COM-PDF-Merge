using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Fábrica y utilidades puras para la resolución y ordenamiento cronológico de ranuras (Fase 1154).
    /// Regla universal canónica: L–V = jornada del aula; SÁBADO = MANANA (08-12) y TARDE (13-17); DOMINGO = MANANA (08-12).
    /// </summary>
    internal static class AgRanuras
    {
        private static readonly Dictionary<string, int> OrdenFranja = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            { AgFranjaJornada.MANANA, 1 },
            { AgFranjaJornada.TARDE, 2 },
            { AgFranjaJornada.NOCHE, 3 }
        };

        /// <summary>
        /// Genera la lista ordenada de ranuras para una ventana dada, jornada del aula y catálogo de bloques activos (Fase 1154).
        /// </summary>
        public static IReadOnlyList<AgRanura> PorVentana(
            IReadOnlyList<int> ventana,
            IReadOnlyCollection<AgBloqueHorario> bloques = null,
            string jornada = null,
            double unidad = 4.0)
        {
            if (ventana == null || ventana.Count == 0) return new List<AgRanura>();

            double unidadEfectiva = unidad > 0 ? unidad : 4.0;
            var diasOrdenados = ventana.OrderBy(d => d).ToList();
            var candidatas = new List<AgRanura>();

            foreach (var dia in diasOrdenados)
            {
                if (dia >= 1 && dia <= 5)
                {
                    var fj = AgFranjaJornada.Resolve(jornada, (int)Math.Round(unidadEfectiva), dia);
                    string franja = fj.HasValue && !string.IsNullOrWhiteSpace(fj.Value.Franja) ? fj.Value.Franja : null;

                    if (franja == null)
                    {
                        if (bloques == null || bloques.Count == 0)
                        {
                            // Fase 1171: sin jornada resoluble ni catálogo no se puede decidir: compatibilidad NOCHE
                            candidatas.Add(new AgRanura(dia, AgFranjaJornada.NOCHE, unidadEfectiva));
                            continue;
                        }

                        // F1179: preferir el bloque con la hora canónica de la jornada antes que el más temprano
                        var delDia = bloques
                            .Where(b => b.Activo == 1 && b.DiaSemanaNumero == dia
                                        && Math.Abs(b.DuracionHoras - unidadEfectiva) < 0.001)
                            .OrderByDescending(b => fj.HasValue
                                && string.Equals(b.HoraInicio, fj.Value.HoraInicio, StringComparison.OrdinalIgnoreCase))
                            .ThenBy(b => b.HoraInicio)
                            .FirstOrDefault();
                        if (delDia == null) continue;   // catálogo confirma ausencia de ranura (I1_CONF_RANURA)
                        franja = delDia.Franja;
                    }

                    candidatas.Add(new AgRanura(dia, franja, unidadEfectiva));
                }
                else if (dia == 6)
                {
                    // Fase 1165: Sábado con 8h es 1 ranura (08:00-17:00); con 4h son 2 ranuras (MANANA 08-12 y TARDE 13-17)
                    if (unidadEfectiva >= 8.0)
                    {
                        candidatas.Add(new AgRanura(6, AgFranjaJornada.MANANA, unidadEfectiva));
                    }
                    else
                    {
                        candidatas.Add(new AgRanura(6, AgFranjaJornada.MANANA, unidadEfectiva));
                        candidatas.Add(new AgRanura(6, AgFranjaJornada.TARDE, unidadEfectiva));
                    }
                }
                else if (dia == 7)
                {
                    // Domingo: 1 ranura canónica (MANANA 08-12 o 08-17)
                    candidatas.Add(new AgRanura(7, AgFranjaJornada.MANANA, unidadEfectiva));
                }
            }

            // Si se proporciona catálogo de bloques, filtrar ranuras que tengan bloque activo de 4h
            if (bloques != null && bloques.Count > 0)
            {
                var resultado = new List<AgRanura>();
                foreach (var ranura in candidatas)
                {
                    if (ranura.Dia == 7)
                    {
                        resultado.Add(ranura);   // F1180: DOM sin catálogo por diseño (V72); no descartar
                        continue;
                    }

                    var permitidas = FranjasDe(ranura.Dia, jornada, ranura.Horas);
                    bool existeBloque = bloques.Any(b =>
                        b.Activo == 1
                        && b.DiaSemanaNumero == ranura.Dia
                        && Math.Abs(b.DuracionHoras - ranura.Horas) < 0.001
                        && (permitidas.Count == 0
                            ? string.Equals(b.Franja, ranura.Franja, StringComparison.OrdinalIgnoreCase)
                            : permitidas.Any(f => string.Equals(f, b.Franja, StringComparison.OrdinalIgnoreCase))));

                    if (existeBloque)
                    {
                        resultado.Add(ranura);
                    }
                }

                if (resultado.Count == 0)
                {
                    // F1189: reparar por día con el catálogo real; nunca devolver ranuras sin bloque
                    foreach (var dia in diasOrdenados)
                    {
                        var delDia = bloques
                            .Where(b => b.Activo == 1 && b.DiaSemanaNumero == dia
                                        && Math.Abs(b.DuracionHoras - unidadEfectiva) < 0.001)
                            .GroupBy(b => (b.Franja ?? string.Empty).Trim().ToUpperInvariant())
                            .Select(g => g.OrderBy(b => b.HoraInicio).First())
                            .OrderBy(b => b.HoraInicio)
                            .ToList();

                        int tope = (dia == 6 && unidadEfectiva < 8.0) ? 2 : 1;
                        foreach (var b in delDia.Take(tope))
                        {
                            resultado.Add(new AgRanura(dia, b.Franja, unidadEfectiva));
                        }
                    }

                    if (resultado.Count > 0)
                    {
                        LoggerHelper.WarningLimitado(
                            $"RanurasReparadasCatalogo|{string.Join(",", diasOrdenados)}|{unidadEfectiva}|{jornada}",
                            $"[Ranuras] Ventana '{string.Join(",", diasOrdenados)}' sin ranuras canónicas de {unidadEfectiva:0.#}h " +
                            $"(jornada '{jornada}'); se reparó con catálogo: [{string.Join(", ", resultado.Select(r => r.ToString()))}].");
                    }
                }

                return Ordenar(resultado);
            }

            return Ordenar(candidatas);
        }

        /// <summary>
        /// Sobrecarga textual para ventana en formato "JUE-SAB", "VIE,SAB,DOM", etc.
        /// </summary>
        public static IReadOnlyList<AgRanura> PorVentana(
            string diasVentana,
            IReadOnlyCollection<AgBloqueHorario> bloques = null,
            string jornada = null,
            double unidad = 4.0)
        {
            var ventana = AgDiasVentana.Dias(diasVentana);
            return PorVentana(ventana, bloques, jornada, unidad);
        }

        /// <summary>
        /// Ordena cronológicamente las ranuras por Día y luego por Franja (MANANA -> TARDE -> NOCHE).
        /// </summary>
        public static IReadOnlyList<AgRanura> Ordenar(IEnumerable<AgRanura> ranuras)
        {
            if (ranuras == null) return new List<AgRanura>();
            return ranuras
                .OrderBy(r => r.Dia)
                .ThenBy(r => OrdenFranja.TryGetValue(r.Franja, out var o) ? o : 99)
                .ToList();
        }

        /// <summary>
        /// Resuelve las franjas permitidas para un día según la jornada y duración de sesión (Fase 1166).
        /// </summary>
        public static IReadOnlyList<string> FranjasDe(int dia, string jornada, double unidad = 4.0)
        {
            double u = unidad > 0 ? unidad : 4.0;
            if (dia >= 1 && dia <= 5)
            {
                // F1186: incluir franjas equivalentes por ventana horaria (NOCHE 8h ≡ TARDE 8h).
                // F1172: jornada no resoluble -> sin restricción (R11 se omite).
                return AgFranjaJornada.Equivalentes(jornada, (int)Math.Round(u), dia);
            }
            if (dia == 6)
            {
                if (u >= 8.0) return new[] { AgFranjaJornada.MANANA };
                return new[] { AgFranjaJornada.MANANA, AgFranjaJornada.TARDE };
            }
            if (dia == 7)
            {
                return new[] { AgFranjaJornada.MANANA };
            }
            return new string[0];
        }
    }
}
