using System;

namespace ConversorPDF.XLL.Core.Models.Academic
{
    public class Sesion
    {
        public int Id { get; set; }
        public int IdVinculacion { get; set; }
        public int IdTercero { get; set; }
        public DateTime fechasesionInicial { get; set; }
        public DateTime fechasesionFinal { get; set; }
        public double Duracion { get; set; }
        public double Horas { get; set; }
        public bool Viatica { get; set; }
    }
}
