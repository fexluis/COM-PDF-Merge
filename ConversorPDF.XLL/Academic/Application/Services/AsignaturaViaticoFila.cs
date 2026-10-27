using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Application.Services
{
    /// <summary>
    /// Representa una fila resultante de exportación de asignaturas a Excel,
    /// conteniendo la definición de la asignatura base y los bloques de viáticos asignados (máx. 8).
    /// </summary>
    public sealed class AsignaturaViaticoFila
    {
        /// <summary>
        /// Sesiones que componen la asignatura/sede/ventana base de esta fila.
        /// </summary>
        public IReadOnlyList<SesionListItem> Asignatura { get; set; }

        /// <summary>
        /// Bloques de viáticos asignados a esta fila (máximo 8 comisiones).
        /// </summary>
        public List<ViaticoBloque> Bloques { get; } = new List<ViaticoBloque>();

        /// <summary>
        /// Indica si esta fila corresponde a una fila de continuación (último recurso R5).
        /// </summary>
        public bool EsContinuacion { get; set; }

        /// <summary>
        /// Contador de bloques asignados a la fila.
        /// </summary>
        public int Contador => Bloques.Count;
    }
}
