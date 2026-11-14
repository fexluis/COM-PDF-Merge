using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>Búsqueda acotada de sesiones reales por bandas; permite comparar y compactar planes factibles.</summary>
    internal static class AgPlanificadorRanuras
    {
        internal static IReadOnlyList<AgSesionPlanificada> Materializar(int id, IReadOnlyList<AgRanura> ranuras,
            int sesiones, DateTime ini, DateTime fin, Func<DateTime, bool> festivo)
        {
            var calendario = Calendario(ranuras, ini, fin, festivo, CancellationToken.None);
            if (sesiones <= 0 || sesiones > calendario.Count) { return new List<AgSesionPlanificada>(); }
            var disponibles = ranuras.Select(r => calendario.Where(s => s.Ranura.Equals(r)).ToList()).ToList();
            int restante = sesiones;
            var resultado = new List<AgSesionPlanificada>();
            for (int i = 0; i < ranuras.Count; i++)
            {
                int objetivo = (int)Math.Ceiling(restante / (double)(ranuras.Count - i));
                int otras = disponibles.Skip(i + 1).Sum(x => x.Count);
                int tomar = Math.Max(Math.Min(objetivo, disponibles[i].Count), restante - otras);
                resultado.AddRange(disponibles[i].Take(tomar).Select(s => new AgSesionPlanificada(id, s.Fecha, s.Ranura)));
                restante -= tomar;
            }
            return resultado.OrderBy(s => s.Fecha).ThenBy(s => s.Ranura.Franja, StringComparer.Ordinal).ToList();
        }

        internal static AgPlanOleadasResultado Construir(
            IReadOnlyList<(int IdAsignatura, double Horas, int OrdenCurricular)> asignaturas,
            IReadOnlyList<AgRanura> ranuras, DateTime ini, DateTime fin, double unidad,
            Func<DateTime, bool> festivo, IReadOnlyDictionary<int, string> patrones,
            int presupuesto = 12000, CancellationToken ct = default(CancellationToken), string diagnostico = null)
        {
            var orden = asignaturas.OrderBy(a => a.OrdenCurricular > 0 ? a.OrdenCurricular : 0)
                .ThenBy(a => a.IdAsignatura).ToList();
            var requeridas = orden.ToDictionary(a => a.IdAsignatura, a => (int)Math.Ceiling(a.Horas / unidad));
            var calendario = Calendario(ranuras, ini, fin, festivo, ct);
            bool sinPatrones = patrones == null || patrones.Count == 0;
            var mejor = new List<AgPlanOleada>();
            var mejorParticion = new List<IReadOnlyList<int>>();
            int mejorSesiones = -1;
            int mejorCompletas = -1;
            DateTime mejorFin = DateTime.MaxValue;
            int nodos = 0;
            bool agotada = false;
            // Una banda permite aprovechar días sobrantes entre materias. Las demás evitan colas de patrones fijos.
            var particiones = new List<List<IReadOnlyList<AgRanura>>>
            {
                new List<IReadOnlyList<AgRanura>> { ranuras }
            };
            particiones.Add(Enumerable.Range(0, AgCarriles.Cantidad(ranuras))
                .Select(c => (IReadOnlyList<AgRanura>)AgCarriles.DelCarril(ranuras, c)).ToList());
            int cortes = Math.Max(0, ranuras.Count - 1);
            for (int mask = 0; mask < (1 << cortes); mask++)
            {
                var bandas = new List<IReadOnlyList<AgRanura>>();
                var banda = new List<AgRanura>();
                for (int i = 0; i < ranuras.Count; i++)
                {
                    banda.Add(ranuras[i]);
                    if (i == ranuras.Count - 1 || (mask & (1 << i)) != 0)
                    {
                        bandas.Add(banda.ToList());
                        banda.Clear();
                    }
                }
                if (bandas.Count > 0) { particiones.Add(bandas); }
            }
            foreach (var bandas in particiones)
            {
                ct.ThrowIfCancellationRequested();
                if (agotada || (sinPatrones && (mejorCompletas == orden.Count || mejorSesiones == calendario.Count))) { break; }
                var calendarios = bandas.Select(b => calendario.Where(s => b.Contains(s.Ranura)).ToList()).ToList();
                var cursores = new int[bandas.Count];
                var oleadas = new int[bandas.Count];
                var planes = new List<AgPlanOleada>();
                var memo = new Dictionary<string, int>();
                Action<int> buscar = null;
                buscar = indice =>
                {
                    ct.ThrowIfCancellationRequested();
                    if (++nodos > presupuesto) { agotada = true; return; }
                    int ubicadas = planes.Sum(p => p.Sesiones.Count);
                    int completas = planes.Count(p => p.Sesiones.Count == requeridas[p.IdAsignatura]);
                    DateTime ultimo = planes.SelectMany(p => p.Sesiones).Select(s => s.Fecha).DefaultIfEmpty(ini.Date).Max();
                    if (ubicadas > mejorSesiones || (ubicadas == mejorSesiones && completas > mejorCompletas)
                        || (ubicadas == mejorSesiones && completas == mejorCompletas && ultimo < mejorFin))
                    {
                        mejor = planes.ToList();
                        mejorSesiones = ubicadas;
                        mejorCompletas = completas;
                        mejorFin = ultimo;
                        mejorParticion = bandas.Select(b => (IReadOnlyList<int>)b.Select(r => r.Dia).Distinct().ToList()).ToList();
                    }
                    if (indice >= orden.Count || agotada || (sinPatrones && (mejorCompletas == orden.Count || mejorSesiones == calendario.Count))) { return; }
                    if (ubicadas + orden.Skip(indice).Sum(a => requeridas[a.IdAsignatura]) < mejorSesiones) { return; }
                    string clave = indice + ":" + string.Join(",", cursores);
                    int previo;
                    if (memo.TryGetValue(clave, out previo) && previo > ubicadas) { return; }
                    memo[clave] = ubicadas;
                    var asig = orden[indice];
                    string texto;
                    var dias = patrones != null && patrones.TryGetValue(asig.IdAsignatura, out texto)
                        ? AgDiasVentana.Dias(texto) : new List<int>();
                    for (int c = 0; c < bandas.Count && !agotada; c++)
                    {
                        if (dias.Any(d => !bandas[c].Any(r => r.Dia == d))) { continue; }
                        var disponibles = calendarios[c].Select((s, i) => new { Sesion = s, Indice = i })
                            .Where(x => x.Indice >= cursores[c] && (dias.Count == 0 || dias.Contains(x.Sesion.Ranura.Dia))).ToList();
                        int tomar = Math.Min(requeridas[asig.IdAsignatura], disponibles.Count);
                        if (tomar <= 0) { continue; }
                        var elegidas = disponibles.Take(tomar).ToList();
                        var sesiones = elegidas.Select(x => new AgSesionPlanificada(asig.IdAsignatura, x.Sesion.Fecha, x.Sesion.Ranura)).ToList();
                        var utilizadas = bandas[c].Where(r => sesiones.Any(s => s.Ranura.Equals(r))).ToList();
                        int anterior = cursores[c];
                        DateTime comienzo = sesiones.First().Fecha;
                        DateTime final = sesiones.Last().Fecha;
                        planes.Add(new AgPlanOleada(asig.IdAsignatura, c, oleadas[c], utilizadas.Select(r => r.Dia).Distinct().ToList(),
                            comienzo, final, (final - comienzo).Days / 7 + 1, requeridas[asig.IdAsignatura], disponibles.Count, utilizadas, sesiones));
                        cursores[c] = elegidas.Last().Indice + 1;
                        oleadas[c]++;
                        buscar(indice + 1);
                        oleadas[c]--;
                        cursores[c] = anterior;
                        planes.RemoveAt(planes.Count - 1);
                    }
                    if (!agotada) { buscar(indice + 1); }
                };
                buscar(0);
            }
            var porId = mejor.ToDictionary(p => p.IdAsignatura, p => p.Sesiones.Count);
            var pendientes = orden.Where(a => !porId.ContainsKey(a.IdAsignatura) || porId[a.IdAsignatura] < requeridas[a.IdAsignatura])
                .Select(a => a.IdAsignatura).ToList();
            if (LoggerHelper.ModoVerboso)
            {
                LoggerHelper.Debug($"[AG-RANURAS-TMP] BUSQUEDA {diagnostico} calendario={calendario.Count} demanda={requeridas.Values.Sum()} nodos={nodos} presupuesto={presupuesto} agotada={agotada} ubicadas={Math.Max(0, mejorSesiones)} completas={mejorCompletas}/{orden.Count} pendientes={string.Join(",", pendientes)} fin={mejor.SelectMany(p => p.Sesiones).Select(s => (DateTime?)s.Fecha).Max():yyyy-MM-dd} bandas={string.Join("|", mejorParticion.Select(b => string.Join(",", b)))}");
            }
            return new AgPlanOleadasResultado(mejor, pendientes,
                pendientes.Select(id => "Asig " + id + ": sesiones pendientes dentro del plazo.").ToList(), mejorParticion,
                ocupacionPorcentaje: calendario.Count > 0 ? 100.0 * Math.Max(0, mejorSesiones) / calendario.Count : 0,
                busquedaAgotada: agotada);
        }

        private static List<AgSesionPlanificada> Calendario(IReadOnlyList<AgRanura> ranuras,
            DateTime ini, DateTime fin, Func<DateTime, bool> festivo, CancellationToken ct)
        {
            var resultado = new List<AgSesionPlanificada>();
            for (DateTime fecha = ini.Date; fecha <= fin.Date;)
            {
                ct.ThrowIfCancellationRequested();
                int dia = fecha.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)fecha.DayOfWeek;
                if (festivo == null || !festivo(fecha))
                {
                    resultado.AddRange(ranuras.Where(r => r.Dia == dia).Select(r => new AgSesionPlanificada(0, fecha, r)));
                }
                if (fecha == fin.Date) { break; }
                fecha = fecha.AddDays(1);
            }
            return resultado.OrderBy(s => s.Fecha).ThenBy(s => s.Ranura.Franja, StringComparer.Ordinal).ToList();
        }
    }
}
