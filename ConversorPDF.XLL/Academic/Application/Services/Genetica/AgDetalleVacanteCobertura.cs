using System;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Estructura de diagnóstico para vacantes encontradas durante la certificación de cobertura.
    /// Separa el detalle de diagnóstico forense del código de retorno estable.
    /// </summary>
    public sealed class AgDetalleVacanteCobertura
    {
        public string ClaveValor { get; set; }
        public int IdAsignatura { get; set; }
        public string MotivoVacante { get; set; }
        public DateTime Fecha { get; set; }
        public double Horas { get; set; }
        public bool Consentida { get; set; }
    }
}
