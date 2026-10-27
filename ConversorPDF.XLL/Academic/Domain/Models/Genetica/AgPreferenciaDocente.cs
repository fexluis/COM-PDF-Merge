using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Representa la preferencia horaria declarada por un docente para un día y franja horaria.
    /// Mapea a la tabla 'ag_preferencia_docente'.
    /// </summary>
    public class AgPreferenciaDocente
    {
        public int Id { get; set; }
        public int IdTercero { get; set; }
        public int DiaSemana { get; set; }
        public string Franja { get; set; }
        public int Preferencia { get; set; }
        public string Subunidad { get; set; }

        public AgPreferenciaDocente()
        {
            Franja = "MANANA";
            Preferencia = 1;
            Subunidad = string.Empty;
        }

        public override string ToString() => $"Preferencia[Doc:{IdTercero}, Dia:{DiaSemana}, Franja:{Franja}, Nivel:{Preferencia}]";
    }
}
