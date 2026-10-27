using System;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Servicio utilitario para el recálculo y consolidación de métricas de individuos cromosómicos.
    /// </summary>
    public static class AgMetricas
    {
        public static void Recalcular(AgIndividuo individuo, AgContexto contexto = null)
        {
            if (individuo == null) return;

            int vacantes = 0;
            int fueraVentana = 0;

            if (individuo.Genes != null)
            {
                for (int i = 0; i < individuo.Genes.Count; i++)
                {
                    var g = individuo.Genes[i];
                    if (g.EsVacante == 1 || g.IdTercero <= 0)
                    {
                        vacantes++;
                    }

                    if (contexto != null && contexto.Ventana != null && g.FechaInicioEstimada.HasValue && g.FechaFinEstimada.HasValue)
                    {
                        bool dentro = contexto.Ventana.ContieneRango(g.FechaInicioEstimada.Value, g.FechaFinEstimada.Value);
                        g.DentroDeVentana = dentro ? 1 : 0;
                        if (!dentro) fueraVentana++;
                    }
                }
            }

            if (individuo.HardViolations == 0 && vacantes == 0 && fueraVentana == 0)
            {
                individuo.Estado = AgEstadoIndividuo.Factible;
            }
            else
            {
                individuo.Estado = AgEstadoIndividuo.Infactible;
            }
        }
    }
}
