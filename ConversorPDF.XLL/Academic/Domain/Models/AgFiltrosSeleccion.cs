using System.Collections.Generic;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Contenedor de IDs seleccionados para los filtros del generador de horarios.
    /// </summary>
    public sealed class AgFiltrosSeleccion
    {
        public List<int> PeriodosIds { get; set; } = new List<int>();
        public List<int> SedesIds { get; set; } = new List<int>();
        public List<int> ProgramasIds { get; set; } = new List<int>();
        public string TipoPrograma { get; set; } = AgTipoProgramaCatalog.Todos;
        public List<int> TercerosIds { get; set; } = new List<int>();
    }
}
