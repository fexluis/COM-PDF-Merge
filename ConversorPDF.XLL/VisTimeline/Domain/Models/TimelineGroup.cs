using System.Collections.Generic;

namespace ConversorPDF.XLL.VisTimeline.Domain.Models
{
    /// <summary>
    /// Representa un grupo (fila) en vis-timeline.
    /// Cada sede es un grupo independiente.
    /// </summary>
    public class TimelineGroup
    {
        /// <summary>Identificador del grupo = SedeCiudad.</summary>
        public string id { get; set; }

        /// <summary>Etiqueta visible en el panel izquierdo. Acepta HTML.</summary>
        public string content { get; set; }

        /// <summary>
        /// Clase CSS para colorear el encabezado del grupo.
        /// Ej: "grupo-sede-cali"
        /// </summary>
        public string className { get; set; }

        /// <summary>IDs de los grupos hijos (sedes bajo un docente).</summary>
        public List<string> nestedGroups { get; set; }

        /// <summary>Expandir hijos por defecto.</summary>
        public bool? showNested { get; set; }

        /// <summary>1 = Docente (padre), 2 = Sede (hijo).</summary>
        public int? treeLevel { get; set; }
    }
}
