using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Entidad que representa la excepción (salvedad) manual que hace un usuario 
    /// sobre un conflicto de horario entre dos sesiones para un docente.
    /// </summary>
    public class SalvedadConflicto
    {
        public int Id { get; set; }
        
        /// <summary>
        /// ID de la primera sesión. Por convención, siempre debe ser el ID menor de las dos.
        /// </summary>
        public int IdSesionA { get; set; }
        
        /// <summary>
        /// ID de la segunda sesión. Por convención, siempre debe ser el ID mayor de las dos.
        /// </summary>
        public int IdSesionB { get; set; }
        
        public int IdTercero { get; set; }
        public int IdPeriodo { get; set; }
        public string Justificacion { get; set; }
        public DateTime FechaCreacion { get; set; }
    }
}
