using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public class ConfiguracionDia
    {
        public DayOfWeek DiaSemana { get; set; }
        public int Horas { get; set; }
        public string Nombre { get; set; }
        public string Horario { get; set; }
    }
}
