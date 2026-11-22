using System;
using System.Collections.Generic;
using System.Linq;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Modelo canónico de concurrencia R5 compartido entre evaluador y validador (Fase 2304).
    /// Calcula las asignaturas concurrentes máximas por cohorte respecto a los carriles disponibles.
    /// </summary>
    internal static class AgConcurrenciaCohorte
    {
        internal static AgResultadoConcurrencia Calcular(
            IEnumerable<(string Cohorte, int IdAsignatura, DateTime Ini, DateTime Fin)> intervalos,
            int maxCarriles)
        {
            if (intervalos == null)
            {
                return new AgResultadoConcurrencia
                {
                    Evaluado = false,
                    MaxConcurrentes = 0,
                    CarrilesDisponibles = maxCarriles,
                    Factible = true
                };
            }

            var lista = intervalos.ToList();
            var cohortes = lista
                .Select(x => x.Cohorte ?? string.Empty)
                .Distinct(StringComparer.Ordinal)
                .ToList();
            if (cohortes.Count > 1)
            {
                throw new ArgumentException("Calcular espera una sola cohorte; para evaluar múltiples cohortes con sus propios carriles evalúe cada una por separado.", nameof(intervalos));
            }

            // Fase 2315: Validación de consistencia temporal en intervalos (Fin >= Ini)
            if (lista.Any(x => x.Fin < x.Ini))
            {
                return new AgResultadoConcurrencia
                {
                    Evaluado = false,
                    MaxConcurrentes = 0,
                    CarrilesDisponibles = maxCarriles,
                    Factible = false
                };
            }

            var porCohorte = lista
                .GroupBy(x => x.Cohorte ?? string.Empty, StringComparer.Ordinal)
                .Select(g => g.GroupBy(x => x.IdAsignatura)
                    .Select(a => (Ini: a.Min(x => x.Ini).Date, Fin: a.Max(x => x.Fin).Date))
                    .ToList())
                .ToList();

            int max = 0;
            foreach (var asigs in porCohorte)
            {
                if (asigs.Count == 0) continue;
                var ini = asigs.Min(x => x.Ini);
                var fin = asigs.Max(x => x.Fin);
                for (var d = ini; d <= fin; d = d.AddDays(1))
                {
                    max = Math.Max(max, asigs.Count(x => x.Ini <= d && d <= x.Fin));
                }
            }

            return new AgResultadoConcurrencia
            {
                Evaluado = porCohorte.Count > 0,
                MaxConcurrentes = max,
                CarrilesDisponibles = maxCarriles,
                Factible = porCohorte.Count == 0 || max <= maxCarriles
            };
        }

        internal static IReadOnlyList<(DateTime TramoIni, DateTime TramoFin, int MaxConcurrentes)> DetectarSobrecargas(
            IEnumerable<(int IdAsignatura, DateTime Ini, DateTime Fin)> asignaturas,
            int maxCarriles)
        {
            var sobrecargas = new List<(DateTime, DateTime, int)>();
            if (asignaturas == null) return sobrecargas;

            var asigs = asignaturas
                .GroupBy(x => x.IdAsignatura)
                .Select(a => (Ini: a.Min(x => x.Ini).Date, Fin: a.Max(x => x.Fin).Date))
                .ToList();

            if (asigs.Count == 0) return sobrecargas;

            var minGlobal = asigs.Min(x => x.Ini);
            var maxGlobal = asigs.Max(x => x.Fin);

            DateTime? tramoIni = null;
            DateTime? tramoFin = null;
            int maxSobrecargaTramo = 0;

            for (var dia = minGlobal; dia <= maxGlobal; dia = dia.AddDays(1))
            {
                int activas = asigs.Count(a => a.Ini <= dia && dia <= a.Fin);
                if (activas > maxCarriles)
                {
                    if (!tramoIni.HasValue)
                    {
                        tramoIni = dia;
                        maxSobrecargaTramo = activas;
                    }
                    else if (activas > maxSobrecargaTramo)
                    {
                        maxSobrecargaTramo = activas;
                    }
                    tramoFin = dia;
                }
                else
                {
                    if (tramoIni.HasValue)
                    {
                        sobrecargas.Add((tramoIni.Value, tramoFin.Value, maxSobrecargaTramo));
                        tramoIni = null;
                        tramoFin = null;
                        maxSobrecargaTramo = 0;
                    }
                }
            }

            if (tramoIni.HasValue)
            {
                sobrecargas.Add((tramoIni.Value, tramoFin.Value, maxSobrecargaTramo));
            }

            return sobrecargas;
        }
    }
}
