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
        public int IdTerritorial { get; set; }
        public string Subunidad { get; set; }
        public int IdPeriodo { get; set; }
        public int IdSede { get; set; }
        public int IdPrograma { get; set; }
        public int IdAula { get; set; }
        public int IdAsignatura { get; set; }
        public string Semestre { get; set; }
        public string Grupo { get; set; }
        public string MotivoVacante { get; set; }
        public DateTime Fecha { get; set; }
        public double Horas { get; set; }
        public bool Consentida { get; set; }

        public string IdentidadLegible =>
            $"{IdSede}/{IdPrograma}/{IdAula}/{IdAsignatura}/{Semestre}/{Grupo}";
    }
}
