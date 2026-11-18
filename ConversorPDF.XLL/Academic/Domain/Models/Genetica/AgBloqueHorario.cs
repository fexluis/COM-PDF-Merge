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
        private string _horaInicio;
        private string _horaFin;
        private TimeSpan _timeSpanInicio;
        private TimeSpan _timeSpanFin;
        private string _horarioTexto;

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

        public string HoraInicio
        {
            get => _horaInicio;
            set
            {
                _horaInicio = value;
                if (TimeSpan.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, out var ts))
                {
                    _timeSpanInicio = ts;
                }
                _horarioTexto = $"{_horaInicio} A {_horaFin}";
            }
        }

        public string HoraFin
        {
            get => _horaFin;
            set
            {
                _horaFin = value;
                if (TimeSpan.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, out var ts))
                {
                    _timeSpanFin = ts;
                }
                _horarioTexto = $"{_horaInicio} A {_horaFin}";
            }
        }

        public TimeSpan TimeSpanInicio => _timeSpanInicio;
        public TimeSpan TimeSpanFin => _timeSpanFin;
        public string HorarioTexto => _horarioTexto;

        public double DuracionHoras { get; set; }
        public string Franja { get; set; }
        public int Activo { get; set; }
        public string Subunidad { get; set; }

        public AgBloqueHorario()
        {
            _timeSpanInicio = TimeSpan.FromHours(8);
            _timeSpanFin = TimeSpan.FromHours(12);
            _horaInicio = "08:00";
            _horaFin = "12:00";
            _horarioTexto = "08:00 A 12:00";
            DuracionHoras = 4.0;
            Franja = "MANANA";
            Activo = 1;
            Subunidad = string.Empty;
            DiaSemanaNumero = 1; // Inicializa DiaSemana en DayOfWeek.Monday
        }

        public string ObtenerHorarioTexto()
        {
            return _horarioTexto ?? $"{HoraInicio} A {HoraFin}";
        }

        public override string ToString()
        {
            int idx = (int)DiaSemana;
            string nombreDia = (idx >= 0 && idx < NombresDia.Length) ? NombresDia[idx] : "LUNES";
            return $"{nombreDia} ({HoraInicio} - {HoraFin}, {DuracionHoras}h)";
        }
    }
}
