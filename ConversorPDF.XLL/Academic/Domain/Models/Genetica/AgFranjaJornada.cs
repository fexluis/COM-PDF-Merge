using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Resolver canónico de franjas por jornada y duración en horas.
    /// Garantiza compatibilidad 1:1 con la semilla histórica y decisiones de 4h u 8h continuas.
    /// </summary>
    public static class AgFranjaJornada
    {
        public const string MANANA = "MANANA";
        public const string TARDE = "TARDE";
        public const string NOCHE = "NOCHE";

        /// <summary>
        /// Resuelve la franja canónica para una jornada, número de horas/día y día de la semana (Fase 962 / D6).
        /// D6: Sábado (diaIso=6) con 8 horas impone franja única MAÑANA 08:00 A 17:00.
        /// TARDE 4h = 13:00 A 17:00; TARDE 8h = 14:00 A 22:00; NOCHE 8h = 14:00 A 22:00.
        /// </summary>
        public static FranjaJornada? Resolve(string jornada, int horasDia, int diaIso = 0)
        {
            if (diaIso == 6 && horasDia == 8)
            {
                return new FranjaJornada("08:00", "17:00", MANANA, 8.0);
            }

            bool ocho = horasDia == 8;
            switch (NormalizarJornada(jornada))
            {
                case "MAÑANA":
                    return ocho
                        ? new FranjaJornada("08:00", "17:00", MANANA, 8.0)
                        : new FranjaJornada("08:00", "12:00", MANANA, 4.0);
                case "TARDE":
                    return ocho
                        ? new FranjaJornada("14:00", "22:00", TARDE, 8.0)
                        : new FranjaJornada("13:00", "17:00", TARDE, 4.0);
                case "NOCHE":
                    return ocho
                        ? new FranjaJornada("14:00", "22:00", NOCHE, 8.0)
                        : new FranjaJornada("18:00", "22:00", NOCHE, 4.0);
                default:
                    return null;
            }
        }

        /// <summary>
        /// Franjas equivalentes por ventana horaria canónica (F1186).
        /// NOCHE 8h (14:00-22:00) ≡ TARDE 8h (14:00-22:00): mismo bloque físico en ag_bloque_horario.
        /// </summary>
        public static IReadOnlyList<string> Equivalentes(string jornada, int horasDia, int diaIso = 0)
        {
            var propia = Resolve(jornada, horasDia, diaIso);
            if (!propia.HasValue) return new string[0];

            var lista = new List<string> { propia.Value.Franja };
            foreach (var f in new[] { MANANA, TARDE, NOCHE })
            {
                if (string.Equals(f, propia.Value.Franja, StringComparison.OrdinalIgnoreCase)) continue;
                var cand = Resolve(f, horasDia, diaIso);
                if (!cand.HasValue) continue;
                if (string.Equals(cand.Value.HoraInicio, propia.Value.HoraInicio, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(cand.Value.HoraFin, propia.Value.HoraFin, StringComparison.OrdinalIgnoreCase))
                {
                    lista.Add(f);
                }
            }
            return lista;
        }

        /// <summary>
        /// Normaliza la representación textual de la jornada para el evaluador (Fase 660: renombrado para evitar colisión UDF).
        /// </summary>
        public static string NormalizarJornada(string jornada)
        {
            string j = (jornada ?? string.Empty).Trim().ToUpperInvariant();
            if (j == "MANANA") return "MAÑANA";
            return j;
        }
    }
}
