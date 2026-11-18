using System.Collections.Generic;

namespace ConversorPDF.XLL.Academic.Application.DTOs
{
    /// <summary>
    /// DTO para la verificación y cuadre de sesiones, horas y vacantes de una corrida en staging.
    /// </summary>
    public sealed class CuadreCorrida
    {
        public int Sesiones { get; set; }
        public double Horas { get; set; }
        public int SesionesVacantes { get; set; }
        public int SesionesHuerfanas { get; set; }
        public IReadOnlyList<(int IdPrograma, int Sesiones, double Horas)> PorPrograma { get; set; }
        public bool RequiereCompleto { get; set; }
    }
}
