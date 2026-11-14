using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ConversorPDF.XLL.Academic.Domain.Models;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Evalúa planes fechados con la misma autoridad que la fábrica, dentro del plazo original.</summary>
    internal static class AgEvaluadorDemandaVentana
    {
        internal static double HorasObjetivo(Asignatura asignatura, AgAulaAsignatura relacion)
        {
            double horas = relacion != null && relacion.HorasObjetivo > 0 ? relacion.HorasObjetivo
                : asignatura.IntHoraria > 0 ? asignatura.IntHoraria : asignatura.Horas > 0 ? asignatura.Horas
                : asignatura.Creditos > 0 ? asignatura.Creditos * 16.0 : 64.0;
            double resto = horas % 4.0;
            return resto > 0.001 && resto < 3.999 ? Math.Ceiling(horas / 4.0) * 4.0 : horas;
        }

        internal static List<(int IdAsignatura, double Horas, int OrdenCurricular)> Demandas(AgContexto contexto, AgAula aula)
        {
            var resultado = new List<(int, double, int)>();
            string semestre = AgFormatoAsignatura.SemestreRomano(aula.Semestre);
            foreach (var aa in contexto.AulaAsignaturaPorAulaAsig.Values.Where(a => a.IdAula == aula.Id)
                .OrderBy(a => a.OrdenCurricular).ThenBy(a => a.IdAsignatura))
            {
                var asignatura = contexto.Asignaturas.FirstOrDefault(a => a.Id == aa.IdAsignatura);
                if (asignatura == null || (aula.IdPrograma > 0 && asignatura.IdPrograma > 0 && aula.IdPrograma != asignatura.IdPrograma)) { continue; }
                string semestreAsignatura = AgFormatoAsignatura.SemestreRomano(asignatura.Semestre);
                if (!string.IsNullOrWhiteSpace(semestre) && semestre != "SIN ASIGNAR"
                    && !string.IsNullOrWhiteSpace(asignatura.Semestre) && semestreAsignatura != "SIN ASIGNAR"
                    && !string.Equals(semestre, semestreAsignatura, StringComparison.OrdinalIgnoreCase)) { continue; }
                double horas = HorasObjetivo(asignatura, aa);
                if (horas > 0) { resultado.Add((aa.IdAsignatura, horas, aa.OrdenCurricular)); }
            }
            return resultado;
        }

        internal static bool PasoMejora(AgResultadoEvaluacionVentana antes, AgResultadoEvaluacionVentana despues)
        {
            if (antes.CabeEnPlazo || !string.IsNullOrEmpty(antes.Error) || !string.IsNullOrEmpty(despues.Error)
                || despues.CapacidadTotal <= antes.CapacidadTotal || despues.DemandaPendiente > antes.DemandaPendiente) { return false; }
            bool conserva = antes.Plan == null || despues.Plan != null
                && antes.Plan.Planes.Where(p => !antes.Plan.Infactibles.Contains(p.IdAsignatura))
                    .All(p => despues.Plan.Planes.Any(n => n.IdAsignatura == p.IdAsignatura && n.Sesiones.Count >= p.Sesiones.Count));
            return conserva && (despues.DemandaPendiente < antes.DemandaPendiente
                || despues.MateriasFueraVentana.Count < antes.MateriasFueraVentana.Count);
        }

        public static AgResultadoEvaluacionVentana Evaluar(string diasVentana, string jornada, double unidad,
            AgContexto contexto, AgAula aula, DateTime ini, DateTime fin, CancellationToken ct = default(CancellationToken))
        {
            ct.ThrowIfCancellationRequested();
            var res = new AgResultadoEvaluacionVentana();
            if (contexto == null || aula == null || fin < ini || unidad <= 0 || double.IsNaN(unidad) || double.IsInfinity(unidad))
            {
                res.Error = "I1_ENTRADA_INVALIDA: contexto, fechas o unidad inválidos.";
                res.HaySaturacion = true;
                return res;
            }
            var demandas = Demandas(contexto, aula);
            var requeridas = demandas.ToDictionary(a => a.IdAsignatura, a => (int)Math.Ceiling(a.Horas / unidad));
            res.DemandaTotal = requeridas.Values.Sum();
            var ranuras = AgRanuras.PorVentana(diasVentana, contexto.Bloques, jornada, unidad);
            res.CapacidadTotal = ranuras.Sum(r => AgPlanificadorTemporal.ContarOcurrencias(r.Dia, ini, fin, contexto.EsFestivo));
            var patrones = contexto.AulaAsignaturaPorAulaAsig.Values.Where(a => a.IdAula == aula.Id
                && !string.Equals(a.OrigenPatron, "AUTOMATICO", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(a.DiasPatron)).ToDictionary(a => a.IdAsignatura, a => a.DiasPatron);
            var planificador = new AgPlanificadorOleadas();
            var evaluacion = planificador.ConstruirEvaluacion(demandas, diasVentana, ini, fin, unidad,
                contexto.EsFestivo, contexto.Bloques, jornada, patrones, ct);
            if (evaluacion.IsFailure)
            {
                res.Error = evaluacion.Error;
                res.Motivo = evaluacion.Error;
                res.HaySaturacion = true;
                res.DemandaPendiente = res.DemandaTotal;
                return res;
            }
            res.Plan = evaluacion.Value;
            res.Sesiones = evaluacion.Value.Planes.SelectMany(p => p.Sesiones).ToList();
            var ubicadas = requeridas.ToDictionary(p => p.Key, _ => 0);
            foreach (var sesion in res.Sesiones) { ubicadas[sesion.IdAsignatura]++; }
            res.IdAsignaturasSaturadas = requeridas.Where(p => ubicadas[p.Key] < p.Value).Select(p => p.Key).ToList();
            res.DemandaPendiente = requeridas.Sum(p => Math.Max(0, p.Value - ubicadas[p.Key]));
            res.RanurasOciosas = Math.Max(0, res.CapacidadTotal - res.Sesiones.Count);
            res.CabeEnPlazo = res.DemandaPendiente == 0 && evaluacion.Value.Infactibles.Count == 0;
            res.BusquedaAgotada = evaluacion.Value.BusquedaAgotada;
            res.HaySaturacion = !res.CabeEnPlazo;
            res.SaturacionCarriles = res.DemandaPendiente > Math.Max(0, res.DemandaTotal - res.CapacidadTotal);
            var dias = new HashSet<int>(ranuras.Select(r => r.Dia));
            res.MateriasFueraVentana = patrones.Where(p => requeridas.ContainsKey(p.Key)
                && AgDiasVentana.Dias(p.Value).Any(d => !dias.Contains(d))).Select(p => p.Key).ToList();
            res.Motivo = res.CabeEnPlazo ? "plan fechado factible dentro del plazo"
                : res.BusquedaAgotada ? "búsqueda acotada: quedan sesiones sin ubicar; no se acredita imposibilidad"
                : "sesiones sin ubicar dentro del plazo por calendario o restricciones";
            res.FechaFinProyectada = res.CabeEnPlazo && res.Sesiones.Count > 0
                ? (DateTime?)res.Sesiones.Max(s => s.Fecha) : null;
            if (!res.CabeEnPlazo && ranuras.Count > 0)
            {
                DateTime horizonte = ini.Date <= DateTime.MaxValue.Date.AddYears(-2)
                    ? ini.Date.AddYears(2).AddDays(-1) : DateTime.MaxValue.Date;
                var proyeccion = planificador.ConstruirEvaluacion(demandas, diasVentana, ini, horizonte, unidad,
                    contexto.EsFestivo, contexto.Bloques, jornada, patrones, ct);
                if (proyeccion.IsSuccess && proyeccion.Value.Infactibles.Count == 0)
                {
                    res.FechaFinProyectada = proyeccion.Value.Planes.SelectMany(p => p.Sesiones).Select(s => (DateTime?)s.Fecha).Max();
                }
                res.HorizonteAgotado = !res.FechaFinProyectada.HasValue;
            }
            if (res.FechaFinProyectada.HasValue)
            {
                res.SemanasRequeridas = (res.FechaFinProyectada.Value.Date - ini.Date).Days / 7 + 1;
            }
            return res;
        }
    }
}
