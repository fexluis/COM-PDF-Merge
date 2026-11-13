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

            if (demandaTotal <= 0)
            {
                return res;
            }

            if (demandaTotal > capacidadTotal)
            {
                res.HaySaturacion = true;
                res.Motivo = $"demanda {demandaTotal} sesiones > capacidad global {capacidadTotal} ranuras";
                return res;
            }

            // Evaluar saturación por carril / banda agrupada por días
            int nDias = diasLista.Count;
            int numBandas = nDias == 2 ? 2 : Math.Max(1, (int)Math.Ceiling(nDias / 2.0));
            int tamBanda = (int)Math.Ceiling(nDias / (double)numBandas);

            var bandasDias = new List<List<int>>();
            for (int b = 0; b < numBandas; b++)
            {
                int desde = b * tamBanda;
                bandasDias.Add(diasLista.Skip(desde).Take(Math.Min(tamBanda, nDias - desde)).ToList());
            }

            int K = bandasDias.Count;
            var capacidadesCarril = new int[K];
            var diasCarril = new List<HashSet<int>>(K);

            for (int c = 0; c < K; c++)
            {
                var diasBandaSet = new HashSet<int>(bandasDias[c]);
                var ranurasBanda = ranuras.Where(r => diasBandaSet.Contains(r.Dia)).ToList();
                capacidadesCarril[c] = ranurasBanda.Sum(r => AgPlanificadorTemporal.ContarOcurrencias(r.Dia, ini, fin, contexto.EsFestivo));
                diasCarril.Add(diasBandaSet);
            }

            // Asignar demandas a carriles según compatibilidad fija y factibilidad
            // Ordenar asignaturas de mayor a menor demanda para pruning eficiente
            var ordenadas = demandas.OrderByDescending(d => d.Sesiones).ToList();

            var cargas = new int[K];
            bool cabe = PuedeAcomodar(0, ordenadas, cargas, capacidadesCarril, diasCarril);

            if (!cabe)
            {
                res.HaySaturacion = true;
                res.SaturacionCarriles = true;
                res.Motivo = $"saturación de carril: demanda {demandaTotal} cabe nominalmente ({capacidadTotal}) pero no es particionable en los carriles disponibles";
            }

            return res;
        }

        private static bool PuedeAcomodar(
            int idx,
            List<(int Sesiones, HashSet<int> DiasPatron)> demandas,
            int[] cargas,
            int[] capacidades,
            List<HashSet<int>> diasCarril)
        {
            if (idx >= demandas.Count) return true;

            var item = demandas[idx];
            int s = item.Sesiones;
            var patron = item.DiasPatron;

            for (int c = 0; c < capacidades.Length; c++)
            {
                // Si la asignatura tiene restricción fija de días, verificar compatibilidad
                if (patron != null && patron.Count > 0)
                {
                    if (patron.Any(d => !diasCarril[c].Contains(d)))
                        continue;
                }

                if (cargas[c] + s <= capacidades[c])
                {
                    cargas[c] += s;
                    if (PuedeAcomodar(idx + 1, demandas, cargas, capacidades, diasCarril))
                        return true;
                    cargas[c] -= s;
                }
            }

            return false;
        }

        private static List<(int Sesiones, HashSet<int> DiasPatron)> ObtenerDemandasAsignaturas(
            AgContexto contexto,
            AgAula aula,
            double unidad)
        {
            var resultado = new List<(int, HashSet<int>)>();
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
                    resultado.Add((sesiones, diasPatron));
                }
            }
            return resultado;
        }
    }
}
