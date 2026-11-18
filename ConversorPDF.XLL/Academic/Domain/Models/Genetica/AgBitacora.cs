using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Registro de telemetría y convergencia por generación durante la ejecución del AG.
    /// Mapea a la tabla 'ag_bitacora'.
    /// </summary>
    public class AgBitacora
    {
        public long Id { get; set; }
        public long IdCorrida { get; set; }
        public int Generacion { get; set; }
        public double MejorFitness { get; set; }
        public double PromedioFitness { get; set; }
        public double PeorFitness { get; set; }
        public int HardViolations { get; set; }
        public int Evaluados { get; set; }
        public int Descartados { get; set; }
        public int Vacantes { get; set; }
        public int FueraDeVentana { get; set; }
        public long DuracionMs { get; set; }
        public string FechaRegistro { get; set; }
        public string Subunidad { get; set; }

        public AgBitacora()
        {
            FechaRegistro = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");
            Subunidad = string.Empty;
        }
    }
}
