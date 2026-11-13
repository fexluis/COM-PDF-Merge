using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Evaluador unificado de demanda y saturación por carriles y ranuras (Fase 1737).
    /// Detecta saturaciones estructurales internas cuando la demanda global cabría nominalmente
    /// pero no puede alojarse en los carriles disponibles de la ventana.
    /// </summary>
    internal static class AgEvaluadorDemandaVentana
    {
        public static AgResultadoEvaluacionVentana Evaluar(
            string diasVentana,
            string jornada,
            double unidad,
            AgContexto contexto,
            AgAula aula,
            DateTime ini,
            DateTime fin)
        {
            var res = new AgResultadoEvaluacionVentana();
            if (contexto == null || aula == null || string.IsNullOrWhiteSpace(diasVentana))
            {
                return res;
            }

            var diasLista = AgDiasVentana.Dias(diasVentana);
            if (diasLista.Count == 0)
            {
                res.HaySaturacion = true;
                res.Motivo = "ventana sin días configurados";
                return res;
            }

            var ranuras = AgRanuras.PorVentana(diasLista, contexto.Bloques, jornada, unidad);
            if (ranuras == null || ranuras.Count == 0)
            {
                res.HaySaturacion = true;
                res.Motivo = "ventana sin ranuras activas";
                return res;
            }

            int capacidadTotal = ranuras.Sum(r => AgPlanificadorTemporal.ContarOcurrencias(r.Dia, ini, fin, contexto.EsFestivo));
            res.CapacidadTotal = capacidadTotal;

            // Obtener demandas por asignatura del aula
            var demandas = ObtenerDemandasAsignaturas(contexto, aula, unidad);
            int demandaTotal = demandas.Sum(d => d.Sesiones);
            res.DemandaTotal = demandaTotal;
            res.RanurasOciosas = Math.Max(0, capacidadTotal - demandaTotal);
            res.DemandaPendiente = Math.Max(0, demandaTotal - capacidadTotal);

            if (demandaTotal <= 0)
            {
                return res;
            }

            // Calcular semanas requeridas y fecha fin mínima proyectada sobre carriles
            int numCarriles = AgCarriles.Cantidad(ranuras);
            var carriles = new List<IReadOnlyList<AgRanura>>(numCarriles);
            for (int c = 0; c < numCarriles; c++)
            {
                carriles.Add(AgCarriles.DelCarril(ranuras, c));
            }

            // Fase 1759: Capacidad y carga granular por (Dia, Franja)
            var slotCarga = ranuras.ToDictionary(r => (r.Dia, r.Franja), _ => 0);
            var cargaCarriles = new int[numCarriles];
            var materiasFueraVentana = new List<int>();

            foreach (var d in demandas)
            {
                int idxCarril = ElegirCarril(d, carriles, cargaCarriles);
                if (idxCarril < 0)
                {
                    // Asignatura fija/legada incompatible con los carriles de esta ventana
                    materiasFueraVentana.Add(d.IdAsignatura);
                    continue;
                }

                // Distribuir entre las ranuras compatibles con el patrón del carril
                var slotsPatron = carriles[idxCarril]
                    .Where(r => d.DiasPatron == null || d.DiasPatron.Count == 0 || d.DiasPatron.Contains(r.Dia))
                    .ToList();
                if (slotsPatron.Count == 0)
                {
                    slotsPatron = carriles[idxCarril].ToList();
                }

                int ranurasEnCarril = Math.Max(1, slotsPatron.Count);
                int s = d.Sesiones;
                for (int i = 0; i < ranurasEnCarril; i++)
                {
                    int fraccion = s / ranurasEnCarril + (i < s % ranurasEnCarril ? 1 : 0);
                    cargaCarriles[idxCarril] += fraccion;
                    var claveSlot = (slotsPatron[i].Dia, slotsPatron[i].Franja);
                    if (slotCarga.ContainsKey(claveSlot))
                    {
                        slotCarga[claveSlot] += fraccion;
                    }
                }
            }

            res.SemanasRequeridas = slotCarga.Values.DefaultIfEmpty(0).Max();
            if (slotCarga.Count > 0)
            {
                var cuello = slotCarga.OrderByDescending(kv => kv.Value)
                    .ThenBy(kv => kv.Key.Dia)
                    .ThenBy(kv => kv.Key.Franja)
                    .First();
                res.DiaCuello = cuello.Key.Dia;
                res.FranjaCuello = cuello.Key.Franja;
            }
            res.FechaFinMinima = FinDeSemanaValida(ini, res.SemanasRequeridas, res.DiaCuello, contexto.EsFestivo);

            // Fase 1759: Saturación directa por slot (día, franja) frente al calendario
            bool saturacionPorSlot = slotCarga.Any(kv =>
                kv.Value > AgPlanificadorTemporal.ContarOcurrencias(kv.Key.Dia, ini, fin, contexto.EsFestivo));

            res.MateriasFueraVentana = materiasFueraVentana.ToList();
            if (materiasFueraVentana.Count > 0)
            {
                res.HaySaturacion = true;
                res.SaturacionCarriles = true;
                res.Motivo = $"materias fuera de ventana: {materiasFueraVentana.Count} asignaturas incompatibles con la estructura de carriles";
                res.IdAsignaturasSaturadas = materiasFueraVentana;
                return res;
            }

            // Fase 1765 & Fase 1766: Saturación de franja / capacidad global y aislamiento de asignaturas saturadas
            if (saturacionPorSlot || demandaTotal > capacidadTotal)
            {
                res.HaySaturacion = true;
                if (saturacionPorSlot)
                {
                    res.SaturacionCarriles = true;
                    res.Motivo = "saturación de franja: carga excede capacidad del calendario en franjas específicas";
                    // Fase 1766: Identificar quirúrgicamente las asignaturas que cargan los slots saturados
                    var slotsSaturados = slotCarga
                        .Where(kv => kv.Value > AgPlanificadorTemporal.ContarOcurrencias(kv.Key.Dia, ini, fin, contexto.EsFestivo))
                        .Select(kv => kv.Key)
                        .ToHashSet();

                    res.IdAsignaturasSaturadas = demandas
                        .Where(d =>
                        {
                            if (d.DiasPatron != null && d.DiasPatron.Count > 0)
                            {
                                return slotsSaturados.Any(s => d.DiasPatron.Contains(s.Dia));
                            }
                            return true;
                        })
                        .Select(d => d.IdAsignatura)
                        .Distinct()
                        .ToList();
                }
                else
                {
                    res.Motivo = $"demanda {demandaTotal} sesiones > capacidad global {capacidadTotal} ranuras";
                    // Fase 1770: En déficit global, acotar a asignaturas ordenadas por mayor demanda que saturan la capacidad
                    int acum = 0;
                    res.IdAsignaturasSaturadas = demandas
                        .OrderByDescending(d => d.Sesiones)
                        .TakeWhile(d =>
                        {
                            acum += d.Sesiones;
                            return true; // Conserva todas las ordenadas hasta completar déficit
                        })
                        .Select(d => d.IdAsignatura)
                        .Distinct()
                        .ToList();
                }

                return res;
            }

            return res;
        }

        private static List<(int Sesiones, HashSet<int> DiasPatron, int IdAsignatura, string OrigenPatron)> ObtenerDemandasAsignaturas(
            AgContexto contexto,
            AgAula aula,
            double unidad)
        {
            var resultado = new List<(int, HashSet<int>, int, string)>();
            if (contexto.AulaAsignaturaPorAulaAsig == null) return resultado;

            string semAula = AgFormatoAsignatura.SemestreRomano(aula.Semestre);
            foreach (var kv in contexto.AulaAsignaturaPorAulaAsig)
            {
                if (kv.Key.Item1 != aula.Id || kv.Value == null) continue;
                var asig = contexto.Asignaturas?.FirstOrDefault(a => a.Id == kv.Key.Item2);
                if (asig != null && asig.IdPrograma > 0 && aula.IdPrograma > 0 && asig.IdPrograma != aula.IdPrograma) continue;
                if (asig != null && !string.IsNullOrWhiteSpace(semAula) && semAula != "SIN ASIGNAR" &&
                    !string.Equals(AgFormatoAsignatura.SemestreRomano(asig.Semestre), semAula, StringComparison.OrdinalIgnoreCase)) continue;

                if (kv.Value.HorasObjetivo > 0)
                {
                    int sesiones = (int)Math.Ceiling(kv.Value.HorasObjetivo / unidad);
                    HashSet<int> diasPatron = null;
                    if (!string.IsNullOrWhiteSpace(kv.Value.DiasPatron))
                    {
                        var dias = AgDiasVentana.Dias(kv.Value.DiasPatron);
                        if (dias.Count > 0)
                        {
                            diasPatron = new HashSet<int>(dias);
                        }
                    }
                    resultado.Add((sesiones, diasPatron, kv.Key.Item2, kv.Value.OrigenPatron ?? "AUTOMATICO"));
                }
            }
            return resultado;
        }

        private static int ElegirCarril(
            (int Sesiones, HashSet<int> DiasPatron, int IdAsignatura, string OrigenPatron) d,
            List<IReadOnlyList<AgRanura>> carriles,
            int[] cargaCarriles)
        {
            // Si el patrón es fijo/legado y tiene días especificados, debe coincidir con todos los días del carril
            bool esAutomatico = string.Equals(d.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase);
            if (!esAutomatico && d.DiasPatron != null && d.DiasPatron.Count > 0)
            {
                for (int c = 0; c < carriles.Count; c++)
                {
                    if (d.DiasPatron.All(day => carriles[c].Any(r => r.Dia == day)))
                    {
                        return c;
                    }
                }
                return -1; // Fuera de ventana / incompatible con los carriles formados
            }

            int minIdx = 0;
            int minCarga = int.MaxValue;
            for (int c = 0; c < cargaCarriles.Length; c++)
            {
                if (cargaCarriles[c] < minCarga)
                {
                    minCarga = cargaCarriles[c];
                    minIdx = c;
                }
            }
            return minIdx;
        }

        private static DateTime FinDeSemanaValida(
            DateTime ini,
            int semanasRequeridas,
            int diaCuello,
            Func<DateTime, bool> esFestivo)
        {
            if (semanasRequeridas <= 0) return ini;
            if (diaCuello < 1 || diaCuello > 7) return ini;

            var cursor = ini.Date;
            int validas = 0;
            while (validas < semanasRequeridas)
            {
                int iso = cursor.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)cursor.DayOfWeek;
                if (iso == diaCuello && (esFestivo == null || !esFestivo(cursor)))
                {
                    validas++;
                    if (validas == semanasRequeridas)
                    {
                        return cursor;
                    }
                }
                cursor = cursor.AddDays(1);
            }
            return ini;
        }
    }
}
