using System;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>Sesión inmutable con asignatura, fecha y ranura (día, franja y duración).</summary>
    internal sealed class AgSesionPlanificada
    {
        /// <summary>Asignatura propietaria de la sesión.</summary>
        public int IdAsignatura { get; }
        /// <summary>Fecha del calendario, sin componente de hora.</summary>
        public DateTime Fecha { get; }
        /// <summary>Día ISO, franja y duración de la sesión.</summary>
        public AgRanura Ranura { get; }

        public AgSesionPlanificada(int idAsignatura, DateTime fecha, AgRanura ranura)
        {
            IdAsignatura = idAsignatura;
            Fecha = fecha.Date;
            Ranura = ranura;
        }
    }
}
