using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Franja o bloque horario discreto disponible para la programación académica.
    /// Mapea a la tabla 'ag_bloque_horario'.
    /// </summary>
    public class AgBloqueHorario
    {
        private static readonly string[] NombresDia =
        { "DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO" };

        private int _diaSemanaNumero = 1;

        public int Id { get; set; }

        /// <summary>1=LUNES .. 7=DOMINGO (convención única de ag_bloque_horario.dia_semana).</summary>
        public int DiaSemanaNumero
        {
            get => _diaSemanaNumero;
            set
            {
                _diaSemanaNumero = value;
                DiaSemana = (value >= 1 && value <= 7)
                    ? (value == 7 ? DayOfWeek.Sunday : (DayOfWeek)value)
                    : DayOfWeek.Monday;
            }
        }

        public DayOfWeek DiaSemana { get; private set; }
        public string HoraInicio { get; set; }
        public string HoraFin { get; set; }
        public double DuracionHoras { get; set; }
        public string Franja { get; set; }
        public int Activo { get; set; }
        public string Subunidad { get; set; }

        public AgBloqueHorario()
        {
            HoraInicio = "08:00";
            HoraFin = "12:00";
            DuracionHoras = 4.0;
            Franja = "MANANA";
            Activo = 1;
            Subunidad = string.Empty;
            DiaSemanaNumero = 1; // Inicializa DiaSemana en DayOfWeek.Monday
        }

        public string ObtenerHorarioTexto()
        {
            return $"{HoraInicio} A {HoraFin}";
        }

        public override string ToString()
        {
            int idx = (int)DiaSemana;
            string nombreDia = (idx >= 0 && idx < NombresDia.Length) ? NombresDia[idx] : "LUNES";
            return $"{nombreDia} ({HoraInicio} - {HoraFin}, {DuracionHoras}h)";
        }
    }
}
