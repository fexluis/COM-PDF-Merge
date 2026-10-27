using System;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Entidad de dominio para aulas físicas o virtuales gestionadas por el AG.
    /// Mapea a la tabla 'ag_aula'.
    /// </summary>
    public class AgAula
    {
        public int Id { get; set; }
        public int IdSede { get; set; }
        public int IdPeriodo { get; set; }
        public int IdPrograma { get; set; }
        public string Codigo { get; set; }
        public string Nombre { get; set; }
        public int Capacidad { get; set; }
        public string TipoAula { get; set; }
        public int EsVirtual { get; set; }
        public AgEstadoAula Estado { get; set; }
        public string CaracteristicasJson { get; set; }
        public string Semestre { get; set; }
        public string Grupo { get; set; }
        public string ModalidadVinculacion { get; set; } = ModalidadTutoriaCatalog.Presencial;
        public string FechaInicio { get; set; }
        public string FechaFin { get; set; }
        public string SemestreInicio { get; set; }
        public string SemestreFin { get; set; }
        public string VentanaOrigen { get; set; }
        public string AulaRef { get; set; }
        public string Subunidad { get; set; }
        public string Jornada { get; set; } = string.Empty;
        public int HorasDia { get; set; }
        public string DiasVentana { get; set; } = string.Empty;

        public AgAula()
        {
            Codigo = string.Empty;
            Nombre = string.Empty;
            TipoAula = "ESTANDAR";
            Estado = AgEstadoAula.Disponible;
            CaracteristicasJson = "{}";
            Semestre = string.Empty;
            Grupo = "A";
            FechaInicio = string.Empty;
            FechaFin = string.Empty;
            SemestreInicio = string.Empty;
            SemestreFin = string.Empty;
            VentanaOrigen = string.Empty;
            AulaRef = string.Empty;
            Subunidad = string.Empty;
            Jornada = string.Empty;
            HorasDia = 0;
            DiasVentana = string.Empty;
        }

        public override string ToString() => $"{Codigo} - {Nombre} ({TipoAula}, Cap: {Capacidad})";
    }
}
