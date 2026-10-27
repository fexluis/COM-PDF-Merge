using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Resultado de la planificación por bandas de días (Fase 1092, Fase 1143, Fase 1227).
    /// </summary>
    internal sealed class AgPlanOleadasResultado
    {
        public IReadOnlyList<AgPlanOleada> Planes { get; }
        public IReadOnlyList<int> Infactibles { get; }
        public IReadOnlyList<string> DetalleInfactibles { get; }

        /// <summary>Fase 1227: bandas de días elegidas (días ISO por banda).</summary>
        public IReadOnlyList<IReadOnlyList<int>> Particion { get; }
        /// <summary>Fase 1227: siguiente ventana sugerida por AgEscaleraVentanas ('' si no aplica).</summary>
        public string VentanaSugerida { get; }
        /// <summary>Fase 1227: patrones que la partición obligó a renormalizar.</summary>
        public int PatronesRenormalizados { get; }
        /// <summary>Fase 1227: ocupación del plan en porcentaje (0–100).</summary>
        public double OcupacionPorcentaje { get; }

        public AgPlanOleadasResultado(
            IReadOnlyList<AgPlanOleada> planes,
            IReadOnlyList<int> infactibles,
            IReadOnlyList<string> detalleInfactibles = null,
            IReadOnlyList<IReadOnlyList<int>> particion = null,
            string ventanaSugerida = null,
            int patronesRenormalizados = 0,
            double ocupacionPorcentaje = 0.0)
        {
            Planes = planes ?? new List<AgPlanOleada>();
            Infactibles = infactibles ?? new List<int>();
            DetalleInfactibles = detalleInfactibles ?? new List<string>();
            Particion = particion ?? new List<IReadOnlyList<int>>();
            VentanaSugerida = ventanaSugerida ?? string.Empty;
            PatronesRenormalizados = patronesRenormalizados;
            OcupacionPorcentaje = ocupacionPorcentaje;
        }
    }
}
