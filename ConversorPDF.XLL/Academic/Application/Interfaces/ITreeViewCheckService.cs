using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.DTOs;

namespace ConversorPDF.XLL.Academic.Application.Interfaces
{
    /// <summary>
    /// Contrato para el servicio de expansión de combinaciones seleccionadas en el árbol jerárquico.
    /// </summary>
    public interface ITreeViewCheckService
    {
        /// <summary>
        /// Genera el producto cartesiano de las selecciones realizadas en los 6 niveles jerárquicos.
        /// </summary>
        IReadOnlyCollection<AgContextoSeleccionDto> ExpandirCombinaciones(
            IEnumerable<int> periodosIds,
            IEnumerable<int> sedesIds,
            IEnumerable<int> programasIds,
            IEnumerable<string> semestres,
            IEnumerable<string> grupos,
            IEnumerable<string> modalidades);
    }
}
