using System;

namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// DTO para la visualización y edición desacoplada de períodos en la grilla de gestión académica.
    /// </summary>
    public class PeriodoGridDto
    {
        public int IdPeriodo { get; set; }
        public string PeriodoNombre { get; set; }
        public string PeriodoInicio { get; set; }
        public string PeriodoFin { get; set; }
        public int Activo { get; set; }

        public string EstadoLegible => Activo == 1 ? "Activo" : "Inactivo";
    }
}
