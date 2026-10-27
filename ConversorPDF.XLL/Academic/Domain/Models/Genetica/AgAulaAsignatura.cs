using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Relación de compatibilidad o requerimiento específico entre un aula y una asignatura.
    /// Mapea a la tabla 'ag_aula_asignatura'.
    /// </summary>
    public class AgAulaAsignatura
    {
        public int IdAula { get; set; }
        public int IdAsignatura { get; set; }
        public int Exclusivo { get; set; }
        public int Prioridad { get; set; }
        public double HorasObjetivo { get; set; }
        public int OrdenCurricular { get; set; }
        public int Requerida { get; set; } = 1;
        public string DiasPatron { get; set; } = string.Empty;
        public string Subunidad { get; set; }

        public AgAulaAsignatura()
        {
            Subunidad = string.Empty;
            DiasPatron = string.Empty;
        }
    }
}
