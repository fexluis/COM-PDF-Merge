using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Application.DTOs;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Operaciones masivas sobre el conjunto de docentes seleccionados (Fase 1319 / R7).
    /// </summary>
    internal static class AgTerceroSeleccion
    {
        public static void SeleccionarTodosVisibles(IEnumerable<AgTerceroSeleccionDto> visibles, HashSet<int> seleccionados)
        {
            if (visibles == null || seleccionados == null) return;
            foreach (var t in visibles)
            {
                seleccionados.Add(t.Id);
                t.Seleccionado = true;
            }
        }

        public static void BorrarSeleccionVisibles(IEnumerable<AgTerceroSeleccionDto> visibles, HashSet<int> seleccionados)
        {
            if (visibles == null || seleccionados == null) return;
            foreach (var t in visibles)
            {
                seleccionados.Remove(t.Id);
                t.Seleccionado = false;
            }
        }

        public static int ContarFueraDeFiltro(IEnumerable<int> seleccionados, IEnumerable<AgTerceroSeleccionDto> visibles)
        {
            if (seleccionados == null) return 0;
            var setVisibles = new HashSet<int>((visibles ?? Enumerable.Empty<AgTerceroSeleccionDto>()).Select(v => v.Id));
            return seleccionados.Count(id => !setVisibles.Contains(id));
        }
    }
}
