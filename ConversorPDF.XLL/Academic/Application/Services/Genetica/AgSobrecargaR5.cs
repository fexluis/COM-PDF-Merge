using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Diagnóstico estructurado de sobrecarga R5 para una cohorte en un tramo cronológico (Fase 2325).
    /// </summary>
    internal sealed class AgSobrecargaR5
    {
        public string Cohorte { get; set; } = string.Empty;
        public DateTime TramoIni { get; set; }
        public DateTime TramoFin { get; set; }
        public int MaxConcurrentes { get; set; }
        public int MaxCarriles { get; set; }
        public List<AgGen> GenesCohorte { get; set; } = new List<AgGen>();
    }
}
