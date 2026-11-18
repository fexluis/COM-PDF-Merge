using System;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Interfaces;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Implementación del evaluador de aptitud (Fitness) estandarizado para el motor del AG.
    /// Emplea la función multiobjetivo: F = F1 - F2 + 2*F3 + F4 + 2*F5.
    /// </summary>
    public class AgEvaluadorAptitud : IAgEvaluadorAptitud
    {
        private readonly IAgValidadorReglas _validadorReglas;

        public bool IncluirDetalleJson { get; set; } = false;

        public AgEvaluadorAptitud(IAgValidadorReglas validadorReglas)
        {
            _validadorReglas = validadorReglas ?? throw new ArgumentNullException(nameof(validadorReglas));
        }

        public void EvaluarAptitud(AgIndividuo individuo, AgContexto contexto)
        {
            EvaluarAptitud(individuo, contexto, forzar: false, incluirDetalleJson: IncluirDetalleJson);
        }

        public void EvaluarAptitud(AgIndividuo individuo, AgContexto contexto, bool forzar, bool incluirDetalleJson = false)
        {
            if (individuo == null) throw new ArgumentNullException(nameof(individuo));
            if (contexto == null) throw new ArgumentNullException(nameof(contexto));

            if (!forzar && individuo.Estado != AgEstadoIndividuo.NoEvaluado)
            {
                return; // No reevaluar clones exactos ya evaluados
            }

            if (_validadorReglas is AgValidadorReglas vr)
            {
                vr.Evaluar(individuo, contexto, incluirDetalleJson);
            }
            else
            {
                _validadorReglas.Evaluar(individuo, contexto);
            }

            // F1: Cobertura de cursos programados
            double f1 = individuo.Genes != null ? individuo.Genes.Count : 0.0;

            // F2: Penalización de restricciones duras y blandas
            // F1026 / F940: penalización ponderada por regla, preservando multiplicidad
            var pesos = contexto.Pesos ?? AgPesosFitness.PorDefecto();
            double f2 = individuo.SoftViolations;
            if (individuo.Violaciones != null && individuo.Violaciones.Count > 0)
            {
                int durasRegistradas = 0;
                foreach (var v in individuo.Violaciones)
                {
                    if (!v.EsDura) continue; // las blandas ya vienen ponderadas en SoftViolations
                    f2 += pesos.PorRegla(v.Regla);
                    durasRegistradas++;
                }
                if (individuo.HardViolations > durasRegistradas)
                {
                    f2 += (individuo.HardViolations - durasRegistradas) * pesos.PenalidadHardBase;
                }
            }
            else
            {
                f2 += individuo.HardViolations * pesos.PenalidadHardBase;
            }

            // F3: Recompensa por cobertura temporal y fechas estimadas válidas dentro de la ventana semestral
            double f3 = 0.0;
            if (contexto.Ventana != null && individuo.Genes != null)
            {
                foreach (var g in individuo.Genes)
                {
                    if (g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue &&
                        contexto.Ventana.ContieneRango(g.FechaInicioEstimada.Value, g.FechaFinEstimada.Value))
                    {
                        f3 += 1.0;
                    }
                }
            }

            // F4: Afinidad del pool y cercanía geográfica (residencia del docente == ciudad de la sede)
            double f4 = 0.0;
            if (individuo.Genes != null)
            {
                foreach (var g in individuo.Genes)
                {
                    if (g.IdTercero > 0 &&
                        contexto.DocentesPorId.TryGetValue(g.IdTercero, out var doc) &&
                        contexto.SedesPorId.TryGetValue(g.IdSede, out var sede))
                    {
                        if (!string.IsNullOrWhiteSpace(doc.Residencia) &&
                            !string.IsNullOrWhiteSpace(sede.Ciudad) &&
                            string.Equals(doc.Residencia.Trim(), sede.Ciudad.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            f4 += 1.0;
                        }
                    }
                }
            }

            // F5: Preservación de asignaciones fijas
            double f5 = individuo.Genes != null ? individuo.Genes.Count(g => g.EsFijo == 1) : 0.0;

            // F = F1 - F2 + 2*F3 + F4 + 2*F5
            individuo.Fitness = f1 - f2 + (2.0 * f3) + f4 + (2.0 * f5);

            individuo.Estado = individuo.HardViolations == 0
                ? AgEstadoIndividuo.Factible
                : AgEstadoIndividuo.Infactible;
        }
    }
}
