using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Representa una opción de partición de carriles del catálogo para una ventana de días (Fase 2179).
    /// </summary>
    internal sealed class AgOpcionParticion
    {
        public string Ventana { get; }
        public string Accion { get; }
        public string CarrilesTexto { get; }
        public IReadOnlyList<IReadOnlyList<AgRanura>> Carriles { get; }

        public AgOpcionParticion(
            string ventana,
            string accion,
            string carrilesTexto,
            IReadOnlyList<IReadOnlyList<AgRanura>> carriles)
        {
            Ventana = ventana ?? string.Empty;
            Accion = accion ?? string.Empty;
            CarrilesTexto = carrilesTexto ?? string.Empty;
            Carriles = carriles ?? new List<IReadOnlyList<AgRanura>>();
        }
    }
}
