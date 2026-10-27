using System;

namespace ConversorPDF.XLL.UI
{
    public class DaySessionInfo
    {
        public int Horas { get; set; }
        public string Horario { get; set; }

        public DaySessionInfo(int horas, string horario)
        {
            Horas = horas;
            Horario = horario;
        }
    }
}
