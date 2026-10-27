using System;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Entidad de vinculación académica en la capa de staging del AG.
    /// Mapea a la tabla 'ag_vinculaciones'.
    /// </summary>
    public class AgVinculacion
    {
        public int Id { get; set; }
        public long IdCorrida { get; set; }
        public long IdIndividuo { get; set; }
        public int IdTercero { get; set; }
        public int IdPrograma { get; set; }
        public int IdSede { get; set; }
        public int Idsede { get => IdSede; set => IdSede = value; }
        public int IdAsignatura { get; set; }
        public int IdPeriodo { get; set; }
        public string Semestre { get; set; }
        public string Grupo { get; set; } = string.Empty;
        public string DiaHorasSemanales { get; set; } = string.Empty;
        public string ModalidadVinculacion { get; set; } = ModalidadTutoriaCatalog.Presencial;
        public double TotalHoras { get; set; }
        public string FechaInicio { get; set; }
        public string FechaFin { get; set; }
        public AgEstadoVinculacion Estado { get; set; }
        public int IsTerminalAereo { get; set; }
        public int IdTerminales { get; set; }
        public int IdTerminal { get => IdTerminales; set => IdTerminales = value; }
        public double GastoTerminalAerea { get; set; }
        public int EsVacante { get; set; }
        public string MotivoVacante { get; set; } = string.Empty;
        public int CandidatosPool { get; set; }
        public string Subunidad { get; set; }
        public System.Collections.Generic.List<long> GenesOrigen { get; set; } = new System.Collections.Generic.List<long>();

        public AgVinculacion()
        {
            Semestre = string.Empty;
            Grupo = string.Empty;
            DiaHorasSemanales = string.Empty;
            ModalidadVinculacion = ModalidadTutoriaCatalog.Presencial;
            FechaInicio = string.Empty;
            FechaFin = string.Empty;
            Estado = AgEstadoVinculacion.Borrador;
            IsTerminalAereo = 0;
            IdTerminales = 0;
            GastoTerminalAerea = 0.0;
            EsVacante = 0;
            MotivoVacante = string.Empty;
            CandidatosPool = 0;
            Subunidad = string.Empty;
        }
    }
}
