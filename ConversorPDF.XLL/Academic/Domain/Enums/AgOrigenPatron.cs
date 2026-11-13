namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Procedencia u origen del patrón de días asignado a una asignatura en un aula (Fase 1735).
    /// </summary>
    public enum AgOrigenPatron
    {
        /// <summary>Asignado o recalculado automáticamente por el planificador según demanda.</summary>
        Automatico = 0,

        /// <summary>Fijado explícitamente como restricción fija por el usuario (inmutable).</summary>
        Fijo = 1,

        /// <summary>Patrón preexistente o legado; se trata como fijo hasta reclasificación expresa.</summary>
        Legado = 2
    }
}

