using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Reparador dedicado de secuencia curricular R6 (Fase 2369): mueve la asignatura corta de un
    /// solape de carril después del bloque que la contiene, usando el inventario real del carril.
    /// </summary>
    internal sealed class AgReparadorR6
    {
        public const int MaxIntentos = 12;
        public const int MaxTiempoMs = 2000;

        public static AgIndividuo Reparar(AgIndividuo campeon, AgContexto contexto, IAgEvaluadorAptitud evaluador)
        {
            var res = Ejecutar(campeon, contexto, evaluador);
            return (res != null && res.HuboMejora && res.Individuo != null) ? res.Individuo : campeon;
        }

        public static AgResultadoReparacionR6 Ejecutar(
            AgIndividuo campeon, AgContexto contexto, IAgEvaluadorAptitud evaluadorAptitud)
        {
            var resultado = new AgResultadoReparacionR6
            {
                Individuo = campeon,
                HardPre = campeon?.HardViolations ?? 0,
                HardPost = campeon?.HardViolations ?? 0
            };

            if (campeon?.Genes == null || contexto == null || evaluadorAptitud == null) return resultado;
            if (campeon.HardViolations <= 0) return resultado;

            var sw = Stopwatch.StartNew();
            var actual = campeon;
            int intentos = 0;
            int aceptados = 0;
            var visitados = new HashSet<string>(StringComparer.Ordinal);

            try
            {
                while (aceptados < MaxIntentos && sw.ElapsedMilliseconds < MaxTiempoMs)
                {
                    var pares = DetectarPares(actual);
                    if (pares.Count == 0) break;

                    bool mejoraEnIteracion = false;

                    foreach (var par in pares)
                    {
                        if (aceptados >= MaxIntentos || sw.ElapsedMilliseconds >= MaxTiempoMs) break;

                        var candidatos = GenerarCandidatos(actual, contexto, par);
                        foreach (var cand in candidatos)
                        {
                            intentos++;
                            if (intentos >= MaxIntentos || sw.ElapsedMilliseconds >= MaxTiempoMs) break;

                            if (cand.Vacantes > actual.Vacantes) continue;
                            if (!AgReparadorR5.Validar(cand, actual, contexto)) continue;

                            cand.Estado = AgEstadoIndividuo.NoEvaluado;
                            evaluadorAptitud.EvaluarAptitud(cand, contexto, forzar: true, incluirDetalleJson: true);

                            if (cand.HardViolations >= actual.HardViolations) continue;

                            int hardAntes = actual.HardViolations;
                            actual = cand;
                            aceptados++;
                            mejoraEnIteracion = true;
                            resultado.Detalle.Add(
                                $"[AG-REPARACION-R6] cohorte={par.Cohorte} carril={par.Carril} " +
                                $"victima={par.Victima.Id} hard {hardAntes}->{cand.HardViolations}");
                            break;
                        }

                        if (mejoraEnIteracion) break;
                    }

                    if (!mejoraEnIteracion) break;

                    string estado = ConstruirEstado(actual);
                    if (!visitados.Add(estado)) break;
                }

                resultado.Intentos = intentos;
                resultado.HardPost = actual.HardViolations;
                if (aceptados > 0 && actual.HardViolations < campeon.HardViolations)
                {
                    resultado.HuboMejora = true;
                    resultado.Individuo = actual;
                    LoggerHelper.Info(
                        $"[AG-REPARACION-R6] Hard {resultado.HardPre}->{resultado.HardPost} " +
                        $"intentos={intentos} aceptados={aceptados} tiempo={sw.ElapsedMilliseconds}ms");
                }
                else
                {
                    resultado.HuboMejora = false;
                    resultado.Individuo = campeon;
                    resultado.HardPost = campeon.HardViolations;
                    LoggerHelper.Info(
                        $"[AG-REPARACION-R6] Rollback sin mejora: Hard {campeon.HardViolations} " +
                        $"intentos={intentos} tiempo={sw.ElapsedMilliseconds}ms");
                }

                return resultado;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[AG-REPARACION-R6] omitida: {ex.Message}");
                resultado.HuboMejora = false;
                resultado.Individuo = campeon;
                resultado.HardPost = campeon.HardViolations;
                resultado.Error = ex.Message;
                return resultado;
            }
        }

        private sealed class AsignaturaLane
        {
            public int Id;
            public DateTime Ini;
            public DateTime Fin;
            public int Sesiones;
            public bool EsFijo;
        }

        private sealed class ParR6
        {
            public string Cohorte;
            public string Carril;
            public AsignaturaLane A;
            public AsignaturaLane B;

            public AsignaturaLane Victima
            {
                get
                {
                    if (A.EsFijo && !B.EsFijo) return B;
                    if (B.EsFijo && !A.EsFijo) return A;
                    return A.Sesiones <= B.Sesiones ? A : B;
                }
            }

            public AsignaturaLane Otro => ReferenceEquals(Victima, A) ? B : A;
        }

        private static List<ParR6> DetectarPares(AgIndividuo ind)
        {
            var res = new List<ParR6>();
            if (ind?.Genes == null) return res;

            foreach (var co in ind.Genes.Where(g => g.EsVacante == 0)
                         .GroupBy(g => AgCalendarioConflictos.Cohorte(g)))
            {
                foreach (var lane in co.GroupBy(g => AgCarriles.ClaveCarril(g)))
                {
                    var asigs = lane.GroupBy(g => g.IdAsignatura)
                        .Select(a => new AsignaturaLane
                        {
                            Id = a.Key,
                            Ini = a.Min(x => x.FechaInicioEstimada ?? DateTime.MaxValue),
                            Fin = a.Max(x => x.FechaFinEstimada ?? DateTime.MinValue),
                            Sesiones = a.SelectMany(AgInventarioRanuras.FechasDeGen).Distinct().Count(),
                            EsFijo = a.Any(x => x.EsFijo == 1)
                        })
                        .OrderBy(a => a.Ini)
                        .ToList();

                    for (int i = 0; i < asigs.Count; i++)
                    {
                        for (int j = i + 1; j < asigs.Count; j++)
                        {
                            if (asigs[i].Ini < asigs[j].Fin && asigs[j].Ini < asigs[i].Fin)
                            {
                                res.Add(new ParR6 { Cohorte = co.Key, Carril = lane.Key, A = asigs[i], B = asigs[j] });
                            }
                        }
                    }
                }
            }

            return res;
        }

        private static List<AgIndividuo> GenerarCandidatos(AgIndividuo actual, AgContexto ctx, ParR6 par)
        {
            var res = new List<AgIndividuo>();
            var victima = par.Victima;
            var otro = par.Otro;
            if (victima == null || otro == null || victima.EsFijo) return res;

            var genesV = actual.Genes
                .Where(g => g.EsVacante == 0 && g.IdAsignatura == victima.Id
                         && string.Equals(AgCalendarioConflictos.Cohorte(g), par.Cohorte, StringComparison.Ordinal)
                         && string.Equals(AgCarriles.ClaveCarril(g), par.Carril, StringComparison.Ordinal))
                .ToList();
            if (genesV.Count == 0) return res;

            var inv = AgInventarioRanuras.Construir(actual, ctx, par.Cohorte);
            var carril = inv.Carriles.FirstOrDefault(c => string.Equals(c.Clave, par.Carril, StringComparison.Ordinal));
            if (carril == null || carril.FechasValidas.Count == 0) return res;

            var idxV = new HashSet<int>();
            foreach (var g in genesV)
            {
                int idx = actual.Genes.IndexOf(g);
                if (idx >= 0) idxV.Add(idx);
            }

            int idAulaV = genesV[0].IdAula;
            var ocupadasFuera = new HashSet<DateTime>();
            for (int i = 0; i < actual.Genes.Count; i++)
            {
                if (idxV.Contains(i)) continue;
                var g = actual.Genes[i];
                if (g.EsVacante != 0) continue;
                bool mismaAula = g.IdAula == idAulaV;
                bool mismaCohorte = string.Equals(AgCalendarioConflictos.Cohorte(g), par.Cohorte, StringComparison.Ordinal);
                if (!mismaAula && !mismaCohorte) continue;

                foreach (var f in AgInventarioRanuras.FechasDeGen(g)) ocupadasFuera.Add(f.Date);
            }

            var libres = carril.FechasValidas.Where(f => !ocupadasFuera.Contains(f)).ToList();
            if (libres.Count == 0) return res;

            DateTime inicioMin = otro.Fin.Date.AddDays(1);
            for (int i = 0; i < actual.Genes.Count; i++)
            {
                if (idxV.Contains(i)) continue;
                var g = actual.Genes[i];
                if (g.EsVacante != 0) continue;
                if (!string.Equals(AgCarriles.ClaveCarril(g), par.Carril, StringComparison.Ordinal)) continue;
                if (!g.FechaFinEstimada.HasValue) continue;
                var fin = g.FechaFinEstimada.Value.Date.AddDays(1);
                if (fin > inicioMin) inicioMin = fin;
            }

            int diaV = genesV.Select(g => (g.DiaSemana >= 1 && g.DiaSemana <= 7) ? g.DiaSemana : 1).Min();
            var inicios = libres
                .Where(f => AgInventarioRanuras.Iso(f) == diaV && f >= inicioMin)
                .OrderBy(f => f)
                .Take(6)
                .ToList();

            foreach (var inicio in inicios)
            {
                var clon = actual.Clonar();
                var usadas = new HashSet<DateTime>();
                DateTime cursor = inicio;
                bool exito = true;

                foreach (var g in genesV.OrderBy(x => x.DiaSemana))
                {
                    int idx = actual.Genes.IndexOf(g);
                    if (idx < 0 || idx >= clon.Genes.Count)
                    {
                        exito = false;
                        break;
                    }

                    var gClon = clon.Genes[idx];
                    int n = (gClon.FechasGen != null && gClon.FechasGen.Count > 0)
                        ? gClon.FechasGen.Count
                        : (int)Math.Max(1, Math.Round(gClon.HorasGen / 4.0));

                    var fechas = libres
                        .Where(f => f >= cursor && AgInventarioRanuras.Iso(f) == gClon.DiaSemana && !usadas.Contains(f))
                        .Take(n)
                        .ToList();
                    if (fechas.Count < n)
                    {
                        exito = false;
                        break;
                    }

                    foreach (var f in fechas) usadas.Add(f);
                    gClon.FechasGen = fechas;
                    gClon.FechaInicioEstimada = fechas.Min();
                    gClon.FechaFinEstimada = fechas.Max();
                    cursor = fechas[fechas.Count - 1].AddDays(1);
                }

                if (exito) res.Add(clon);
            }

            return res;
        }

        private static string ConstruirEstado(AgIndividuo ind)
        {
            if (ind?.Genes == null) return string.Empty;
            var sb = new System.Text.StringBuilder(ind.Genes.Count * 20);
            for (int i = 0; i < ind.Genes.Count; i++)
            {
                var g = ind.Genes[i];
                sb.Append(i).Append(':')
                  .Append(g.FechaInicioEstimada.HasValue ? g.FechaInicioEstimada.Value.ToString("yyyyMMdd") : "-").Append(':')
                  .Append(g.FechaFinEstimada.HasValue ? g.FechaFinEstimada.Value.ToString("yyyyMMdd") : "-").Append(';');
            }
            return sb.ToString();
        }
    }
}