using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using CSharpFunctionalExtensions;
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
            var orden = AgOrdenAsignaturas.PorHoras(asignaturas);
            var requeridas = orden.ToDictionary(a => a.IdAsignatura, a => (int)Math.Ceiling(a.Horas / unidad));
            var calendario = Calendario(ranuras, ini, fin, festivo, ct);
            bool sinPatrones = patrones == null || patrones.Count == 0;
            int totalSesiones = requeridas.Values.Sum();
            DateTime cotaMinimaFin = (totalSesiones > 0 && totalSesiones <= calendario.Count)
                ? calendario[totalSesiones - 1].Fecha
                : ini.Date;
            var mejor = new List<AgPlanOleada>();
            var mejorParticion = new List<IReadOnlyList<int>>();
            int mejorSesiones = -1;
            int mejorCompletas = -1;
            DateTime mejorFin = DateTime.MaxValue;
            int nodos = 0;
            bool agotada = false;
            // Una banda permite aprovechar días sobrantes entre materias. Las demás evitan colas de patrones fijos.
            var particiones = new List<List<IReadOnlyList<AgRanura>>>();
            particiones.Add(Enumerable.Range(0, AgCarriles.Cantidad(ranuras))
                .Select(c => (IReadOnlyList<AgRanura>)AgCarriles.DelCarril(ranuras, c)).ToList());
            if (AgCarriles.Cantidad(ranuras) > 1)
            {
                particiones.Add(new List<IReadOnlyList<AgRanura>> { ranuras });
            }
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
                if (agotada || mejorSesiones == calendario.Count) { break; }
                if (sinPatrones && mejorCompletas == orden.Count && mejorFin <= cotaMinimaFin) { break; }
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
                    if (indice >= orden.Count || agotada || (sinPatrones && completas == orden.Count)) { return; }
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
                        int semanas = (final - comienzo).Days / 7 + 1;
                        int cap = utilizadas.Count * semanas;
                        planes.Add(new AgPlanOleada(asig.IdAsignatura, c, oleadas[c], utilizadas.Select(r => r.Dia).Distinct().ToList(),
                            comienzo, final, semanas, requeridas[asig.IdAsignatura], cap, utilizadas, sesiones));
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

        /// <summary>
        /// Empaqueta las cuotas pendientes en los huecos más tempranos posibles respetando los días permitidos
        /// de cada asignatura. Devuelve false si no hay empaquetado factible (el llamador conserva el relleno previo).
        /// </summary>
        private static bool EmpaquetarPendientes(
            IReadOnlyList<(int IdAsignatura, double Horas, int OrdenCurricular)> restantes,
            Dictionary<int, int> cuotas, Dictionary<int, HashSet<int>> restricciones,
            int totalHuecos, Func<int, int> diaDe, Func<int, DateTime> fechaDe, Func<int, AgRanura> ranuraDe,
            bool[] ocupadas, Dictionary<int, List<AgSesionPlanificada>> porId, CancellationToken ct)
        {
            var sujetos = restantes.Where(a => cuotas[a.IdAsignatura] > 0
                && restricciones[a.IdAsignatura].All(d => diaDeParaValidar(diaDe, totalHuecos, d))).ToList();
            // Si no queda nada pendiente no hay trabajo; pero si hay pendientes que el guardián de días
            // válidos excluyó (por ejemplo un día de patrón sin ocurrencia en la ventana), se devuelve
            // false para que el relleno previo siga cubriendo lo que sí puede cubrirse.
            if (sujetos.Count == 0) return !restantes.Any(a => cuotas[a.IdAsignatura] > 0);

            int requerido = sujetos.Sum(a => cuotas[a.IdAsignatura]);
            int libres = 0;
            for (int i = 0; i < totalHuecos; i++) if (!ocupadas[i]) libres++;
            if (libres < requerido) return false;

            int s = 0, baseSujeto = 1, baseHueco = 1 + sujetos.Count;
            int t = baseHueco + totalHuecos;

            // Horizonte mínimo: primer índice donde el flujo consigue colocar todas las cuotas.
            Dictionary<int, int[]> asignado = null;
            for (int lim = 0; lim < totalHuecos && asignado == null; lim++)
            {
                ct.ThrowIfCancellationRequested();
                var red = new RedFlujo(t + 1);
                for (int k = 0; k < sujetos.Count; k++)
                {
                    red.Agregar(s, baseSujeto + k, cuotas[sujetos[k].IdAsignatura]);
                    var dias = restricciones[sujetos[k].IdAsignatura];
                    for (int i = 0; i <= lim; i++)
                    {
                        if (ocupadas[i]) continue;
                        if (dias.Count > 0 && !dias.Contains(diaDe(i))) continue;
                        red.Agregar(baseSujeto + k, baseHueco + i, 1);
                    }
                }
                for (int i = 0; i <= lim; i++)
                {
                    if (!ocupadas[i]) red.Agregar(baseHueco + i, t, 1);
                }
                if (red.MaxFlujo(s, t) != requerido) continue;

                asignado = new Dictionary<int, int[]>();
                for (int k = 0; k < sujetos.Count; k++) asignado[sujetos[k].IdAsignatura] = new int[cuotas[sujetos[k].IdAsignatura]];
                var llenado = new Dictionary<int, int>();
                for (int k = 0; k < sujetos.Count; k++)
                {
                    int id = sujetos[k].IdAsignatura;
                    for (int i = 0; i <= lim; i++)
                    {
                        if (ocupadas[i] || !red.Usada(baseSujeto + k, baseHueco + i)) continue;
                        if (!llenado.ContainsKey(id)) llenado[id] = 0;
                        if (llenado[id] < asignado[id].Length) asignado[id][llenado[id]++] = i;
                    }
                }
            }

            if (asignado == null) return false;

            // Contrato de respaldo: si alguna pendiente no pudo empaquetarse (por ejemplo un día de patrón
            // sin hueco en el calendario), se devuelve false SIN mutar cuotas, ocupadas ni porId, para que el
            // llamador ejecute el relleno previo desde el estado intacto.
            foreach (var a in restantes)
            {
                if (cuotas[a.IdAsignatura] > 0 && !asignado.ContainsKey(a.IdAsignatura)) return false;
            }

            foreach (var a in restantes)
            {
                int[] huecos;
                if (!asignado.TryGetValue(a.IdAsignatura, out huecos)) continue;
                var sesiones = new List<AgSesionPlanificada>();
                foreach (int i in huecos)
                {
                    ocupadas[i] = true;
                    cuotas[a.IdAsignatura]--;
                    sesiones.Add(new AgSesionPlanificada(a.IdAsignatura, fechaDe(i), ranuraDe(i)));
                }
                sesiones = sesiones.OrderBy(x => x.Fecha)
                    .ThenBy(x => x.Ranura.Franja, StringComparer.Ordinal).ToList();
                porId[a.IdAsignatura] = sesiones;
            }
            return true;
        }

        private static bool diaDeParaValidar(Func<int, int> diaDe, int totalHuecos, int dia)
        {
            for (int i = 0; i < totalHuecos; i++) if (diaDe(i) == dia) return true;
            return false;
        }

        /// <summary>Flujo máximo mínimo para el empaquetado de cuotas en huecos.</summary>
        private sealed class RedFlujo
        {
            private readonly int[] _inicio;
            private readonly List<int>[] _adyacencia;
            private readonly List<int> _destino = new List<int>();
            private readonly List<int> _capacidad = new List<int>();
            private readonly int[] _nivel;
            private readonly int[] _iterador;
            private readonly int _nodos;

            internal RedFlujo(int nodos)
            {
                _nodos = nodos;
                _adyacencia = new List<int>[nodos];
                for (int i = 0; i < nodos; i++) _adyacencia[i] = new List<int>();
                _nivel = new int[nodos];
                _iterador = new int[nodos];
                _inicio = new int[nodos];
            }

            internal void Agregar(int desde, int hacia, int capacidad)
            {
                _adyacencia[desde].Add(_destino.Count);
                _destino.Add(hacia);
                _capacidad.Add(capacidad);
                _adyacencia[hacia].Add(_destino.Count);
                _destino.Add(desde);
                _capacidad.Add(0);
            }

            internal bool Usada(int desde, int hacia)
            {
                foreach (int e in _adyacencia[desde])
                {
                    if (_destino[e] == hacia) return _capacidad[e] == 0;
                }
                return false;
            }

            private bool ConstruirNiveles(int origen, int sumidero)
            {
                for (int i = 0; i < _nodos; i++) _nivel[i] = -1;
                var cola = new Queue<int>();
                _nivel[origen] = 0;
                cola.Enqueue(origen);
                while (cola.Count > 0)
                {
                    int u = cola.Dequeue();
                    foreach (int e in _adyacencia[u])
                    {
                        if (_capacidad[e] <= 0 || _nivel[_destino[e]] >= 0) continue;
                        _nivel[_destino[e]] = _nivel[u] + 1;
                        cola.Enqueue(_destino[e]);
                    }
                }
                return _nivel[sumidero] >= 0;
            }

            private int Empujar(int u, int sumidero, int flujo)
            {
                if (u == sumidero) return flujo;
                for (; _iterador[u] < _adyacencia[u].Count; _iterador[u]++)
                {
                    int e = _adyacencia[u][_iterador[u]];
                    int v = _destino[e];
                    if (_capacidad[e] <= 0 || _nivel[v] != _nivel[u] + 1) continue;
                    int enviado = Empujar(v, sumidero, Math.Min(flujo, _capacidad[e]));
                    if (enviado <= 0) continue;
                    _capacidad[e] -= enviado;
                    _capacidad[e ^ 1] += enviado;
                    return enviado;
                }
                return 0;
            }

            internal int MaxFlujo(int origen, int sumidero)
            {
                int total = 0;
                while (ConstruirNiveles(origen, sumidero))
                {
                    for (int i = 0; i < _nodos; i++) _iterador[i] = 0;
                    int enviado;
                    while ((enviado = Empujar(origen, sumidero, int.MaxValue)) > 0) total += enviado;
                }
                return total;
            }
        }

        internal static Result<AgPlanOleadasResultado> CompletarReferencia(
            IReadOnlyList<(int IdAsignatura, double Horas, int OrdenCurricular)> demandas,
            IReadOnlyList<AgRanura> ranuras, DateTime ini, DateTime fin, double unidad,
            Func<DateTime, bool> festivo, IReadOnlyDictionary<int, string> patrones,
            AgPlanOleadasResultado referencia, CancellationToken ct)
        {
            var orden = AgOrdenAsignaturas.PorHoras(demandas);
            var requeridas = orden.ToDictionary(a => a.IdAsignatura,
                a => (int)Math.Ceiling(a.Horas / unidad));
            var restricciones = orden.ToDictionary(a => a.IdAsignatura, a =>
            {
                string texto;
                return new HashSet<int>(patrones != null
                    && patrones.TryGetValue(a.IdAsignatura, out texto)
                    ? AgDiasVentana.Dias(texto) : new List<int>());
            });
            var calendario = Calendario(ranuras, ini.Date, fin.Date, festivo, ct);
            var ocupadas = new bool[calendario.Count];
            var porId = new Dictionary<int, List<AgSesionPlanificada>>();

            if (referencia != null)
            {
                var validacion = ValidarFechado(referencia, demandas, ranuras,
                    ini, fin, unidad, festivo, patrones, ct);
                if (validacion.IsFailure)
                {
                    return Result.Failure<AgPlanOleadasResultado>(validacion.Error);
                }
                foreach (var p in referencia.Planes.Where(p =>
                    restricciones[p.IdAsignatura].Count > 0
                    && !referencia.Infactibles.Contains(p.IdAsignatura)))
                {
                    ct.ThrowIfCancellationRequested();
                    var sesiones = p.Sesiones.OrderBy(s => s.Fecha)
                        .ThenBy(s => s.Ranura.Franja, StringComparer.Ordinal).ToList();
                    foreach (var s in sesiones)
                    {
                        int indice = calendario.FindIndex(c => c.Fecha == s.Fecha
                            && c.Ranura.Equals(s.Ranura));
                        if (indice >= 0)
                        {
                            ocupadas[indice] = true;
                        }
                    }
                    porId[p.IdAsignatura] = sesiones;
                }
            }

            var cuotas = requeridas.ToDictionary(kv => kv.Key,
                kv => porId.ContainsKey(kv.Key) ? 0 : kv.Value);
            var diasDisponibles = new HashSet<int>(ranuras.Select(r => r.Dia));
            var restantes = orden.Where(a => !porId.ContainsKey(a.IdAsignatura)).ToList();

            // Empaquetado con restricciones de día: se asigna cada cuota a los huecos más tempranos posibles
            // resolviendo un flujo máximo sobre los huecos hasta un horizonte mínimo. Si no hay empaquetado
            // factible se conserva el relleno anterior como respaldo.
            bool empaquetado = EmpaquetarPendientes(restantes, cuotas, restricciones,
                calendario.Count, i => calendario[i].Ranura.Dia, i => calendario[i].Fecha,
                i => calendario[i].Ranura, ocupadas, porId, ct);
            if (empaquetado)
            {
                // Sesiones asignadas por el empaquetado; no se ejecuta el relleno por cuotas.
            }
            else
            foreach (var a in restantes)
            {
                if (porId.ContainsKey(a.IdAsignatura)) continue;
                ct.ThrowIfCancellationRequested();
                var sesiones = new List<AgSesionPlanificada>();
                var dias = restricciones[a.IdAsignatura];
                if (dias.Any(d => !ranuras.Any(r => r.Dia == d)))
                {
                    porId[a.IdAsignatura] = sesiones;
                    continue;
                }
                while (cuotas[a.IdAsignatura] > 0)
                {
                    ct.ThrowIfCancellationRequested();
                    int elegido = -1;
                    int primero = -1;
                    for (int i = 0; i < calendario.Count; i++)
                    {
                        ct.ThrowIfCancellationRequested();
                        if (ocupadas[i] || (dias.Count > 0 && !dias.Contains(calendario[i].Ranura.Dia)))
                        {
                            continue;
                        }
                        if (primero < 0) { primero = i; }
                        // La materia actual debe poder completar su cuota en su propio sufijo.
                        int propiasLibres = 0;
                        int propiaCuota = cuotas[a.IdAsignatura] - 1;
                        for (int j = i + 1; j < calendario.Count && propiasLibres < propiaCuota; j++)
                        {
                            if (!ocupadas[j] && (dias.Count == 0 || dias.Contains(calendario[j].Ranura.Dia)))
                            {
                                propiasLibres++;
                            }
                        }
                        bool factible = propiasLibres >= propiaCuota;
                        // Las otras materias pueden aprovechar cualquier hueco que quede libre.
                        foreach (var pendiente in cuotas.Where(kv => kv.Key != a.IdAsignatura && kv.Value > 0))
                        {
                            if (!factible) { break; }
                            var permitidos = restricciones[pendiente.Key];
                            if (permitidos.Any(d => !diasDisponibles.Contains(d))) { continue; }
                            int libres = 0;
                            for (int j = 0; j < calendario.Count && libres < pendiente.Value; j++)
                            {
                                if (j != i && !ocupadas[j]
                                    && (permitidos.Count == 0 || permitidos.Contains(calendario[j].Ranura.Dia)))
                                {
                                    libres++;
                                }
                            }
                            if (libres < pendiente.Value) { factible = false; break; }
                        }
                        if (factible) { elegido = i; break; }
                    }
                    if (elegido < 0) { elegido = primero; }
                    if (elegido < 0) { break; }
                    var hueco = calendario[elegido];
                    ocupadas[elegido] = true;
                    cuotas[a.IdAsignatura]--;
                    sesiones.Add(new AgSesionPlanificada(a.IdAsignatura, hueco.Fecha, hueco.Ranura));
                }
                sesiones = sesiones.OrderBy(s => s.Fecha)
                    .ThenBy(s => s.Ranura.Franja, StringComparer.Ordinal).ToList();
                porId[a.IdAsignatura] = sesiones;
            }

            var carrilPorId = new Dictionary<int, int>();
            if (referencia != null)
            {
                foreach (var p in referencia.Planes)
                {
                    carrilPorId[p.IdAsignatura] = p.Carril;
                }
            }

            var planes = new List<AgPlanOleada>();
            foreach (var a in orden)
            {
                List<AgSesionPlanificada> sesiones;
                if (!porId.TryGetValue(a.IdAsignatura, out sesiones) || sesiones.Count == 0) { continue; }
                var usadas = ranuras.Where(r => sesiones.Any(s => s.Ranura.Equals(r))).ToList();
                DateTime primera = sesiones.First().Fecha;
                DateTime ultima = sesiones.Last().Fecha;
                int capacidad = calendario.Count(c => c.Fecha >= primera && c.Fecha <= ultima
                    && usadas.Any(r => r.Equals(c.Ranura)));
                int carril = carrilPorId.TryGetValue(a.IdAsignatura, out var cId) ? cId : 0;
                planes.Add(new AgPlanOleada(a.IdAsignatura, carril, planes.Count,
                    usadas.Select(r => r.Dia).Distinct().ToList(), primera, ultima,
                    (ultima - primera).Days / 7 + 1, requeridas[a.IdAsignatura], capacidad,
                    usadas, sesiones));
            }
            var pendientes = orden.Where(a => !porId.ContainsKey(a.IdAsignatura) || porId[a.IdAsignatura].Count < requeridas[a.IdAsignatura])
                .Select(a => a.IdAsignatura).ToList();
            var detalle = new List<string>();
            foreach (var id in pendientes)
            {
                if (referencia != null && referencia.Infactibles.Contains(id))
                {
                    var dRef = referencia.DetalleInfactibles.FirstOrDefault(x => x.Contains("Asig " + id));
                    if (!string.IsNullOrEmpty(dRef))
                    {
                        detalle.Add(dRef);
                        continue;
                    }
                }
                detalle.Add("Asig " + id + ": sesiones no ubicadas; el completado no acredita imposibilidad.");
            }
            string ventanaSugerida = (referencia != null && !string.IsNullOrEmpty(referencia.VentanaSugerida))
                ? referencia.VentanaSugerida
                : (pendientes.Count > 0 ? AgEscaleraVentanas.Siguiente(string.Join(",", ranuras.Select(r => r.Dia).Distinct())) : string.Empty);

            var resultado = new AgPlanOleadasResultado(
                planes,
                pendientes,
                detalle,
                new List<IReadOnlyList<int>> { ranuras.Select(r => r.Dia).Distinct().ToList() },
                ventanaSugerida,
                0,
                calendario.Count == 0 ? 0.0 : 100.0 * planes.Sum(p => p.Sesiones.Count) / calendario.Count);
            var comprobacion = ValidarFechado(resultado, demandas, ranuras,
                ini, fin, unidad, festivo, patrones, ct);
            return comprobacion.IsSuccess ? Result.Success(resultado)
                : Result.Failure<AgPlanOleadasResultado>(comprobacion.Error);
        }

        internal static Result ValidarFechado(
            AgPlanOleadasResultado plan,
            IReadOnlyList<(int IdAsignatura, double Horas, int OrdenCurricular)> demandas,
            IReadOnlyList<AgRanura> ranuras, DateTime ini, DateTime fin, double unidad,
            Func<DateTime, bool> festivo, IReadOnlyDictionary<int, string> patrones,
            CancellationToken ct)
        {
            if (plan == null) { return Result.Failure("I5_PLAN_NULO"); }
            if (patrones != null)
            {
                foreach (var d in demandas)
                {
                    string textoPatron;
                    if (patrones.TryGetValue(d.IdAsignatura, out textoPatron)
                        && !string.IsNullOrWhiteSpace(textoPatron)
                        && AgDiasVentana.Dias(textoPatron).Count == 0)
                    {
                        return Result.Failure("I1_PATRON_INVALIDO");
                    }
                }
            }
            var requeridas = demandas.ToDictionary(a => a.IdAsignatura,
                a => (int)Math.Ceiling(a.Horas / unidad));
            var calendario = Calendario(ranuras, ini.Date, fin.Date, festivo, ct);
            var ocupadas = new HashSet<int>();
            var cantidades = new Dictionary<int, int>();
            foreach (var p in plan.Planes)
            {
                ct.ThrowIfCancellationRequested();
                int requerida;
                if (!requeridas.TryGetValue(p.IdAsignatura, out requerida)
                    || cantidades.ContainsKey(p.IdAsignatura)
                    || p.SesionesRequeridas != requerida || p.Sesiones.Count > requerida)
                {
                    return Result.Failure("I5_PLAN_DEMANDA_INCONSISTENTE");
                }
                cantidades[p.IdAsignatura] = p.Sesiones.Count;
                string texto;
                var dias = patrones != null && patrones.TryGetValue(p.IdAsignatura, out texto)
                    ? AgDiasVentana.Dias(texto) : new List<int>();
                if (patrones != null && patrones.TryGetValue(p.IdAsignatura, out texto)
                    && !string.IsNullOrWhiteSpace(texto) && dias.Count == 0)
                {
                    return Result.Failure("I1_PATRON_INVALIDO");
                }
                if (dias.Any(d => !ranuras.Any(r => r.Dia == d)))
                {
                    return Result.Failure("I5_PLAN_SESION_INVALIDA");
                }
                foreach (var s in p.Sesiones)
                {
                    ct.ThrowIfCancellationRequested();
                    if (s.IdAsignatura != p.IdAsignatura || s.Ranura == null
                        || s.Fecha != s.Fecha.Date
                        || Math.Abs(s.Ranura.Horas - unidad) > 0.001
                        || (dias.Count > 0 && !dias.Contains(s.Ranura.Dia))
                        || s.Fecha < p.FechaInicio.Date || s.Fecha > p.FechaFin.Date
                        || !p.RanurasPatron.Any(r => r.Equals(s.Ranura)))
                    {
                        return Result.Failure("I5_PLAN_SESION_INVALIDA");
                    }
                    int indice = calendario.FindIndex(c => c.Fecha == s.Fecha
                        && c.Ranura.Equals(s.Ranura));
                    if (indice < 0 || !ocupadas.Add(indice))
                    {
                        return Result.Failure("I5_PLAN_FECHA_O_COLISION");
                    }
                }
            }
            var pendientes = new HashSet<int>(requeridas.Where(kv =>
                !cantidades.ContainsKey(kv.Key) || cantidades[kv.Key] < kv.Value)
                .Select(kv => kv.Key));
            return pendientes.SetEquals(plan.Infactibles)
                ? Result.Success() : Result.Failure("I5_PLAN_PENDIENTES_INCONSISTENTES");
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
