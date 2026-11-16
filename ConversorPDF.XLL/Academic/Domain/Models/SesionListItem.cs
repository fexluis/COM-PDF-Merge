using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Modelo para representar una fila desnormalizada de sesión en el listado principal.
    /// </summary>
    public class SesionListItem
    {
        public int IdSesion { get; set; }
        public int IdVinculacion { get; set; }
        public long IdCorrida { get; set; }
        public int IdGen { get; set; }
        public int IdAula { get; set; }
        public int EsVacante { get; set; }
        public string MotivoVacante { get; set; }
        public int CandidatosPool { get; set; }
        public int IdPeriodo { get; set; }
        public int IdPrograma { get; set; }
        public string Subunidad { get; set; }
        public string ProgramaTipo { get; set; }
        public string ProgramaAbreviatura { get; set; }
        public string Semestre { get; set; }
        public int? Nivel { get; set; }
        public string DocenteNombre { get; set; }
        public int IdTercero { get; set; }
        public int IdAsignatura { get; set; }
        public string CedulaTercero { get; set; }
        public string AsignaturaNombre { get; set; }
        public string AsignaturaAbreviatura { get; set; }
        public string SedeCiudad { get; set; }
        public string Subsede { get; set; }
        public string DiaSemana { get; set; }
        public double Horas { get; set; }
        /// <summary>
        /// Fecha de la sesión (columna fechasesion en DB, mapeada desde vw_SesionesDetalle.FechaSesion).
        /// </summary>
        public DateTime? FechaSesion { get; set; }
        /// <summary>
        /// Horario descriptivo, ej: "MARTES|4" (columna horariosesion en DB).
        /// </summary>
        public string HorarioSesion { get; set; }
        public bool Viatica { get; set; }
        public string SedeDepartamento { get; set; }
        public string ProgramaNombre { get; set; }
        public int? AsignaturaCreditos { get; set; }
        public string AsignaturaMetodologia { get; set; }
        public string AsignaturaCodMateria { get; set; }
        public string Grupo { get; set; }
        public string VinculacionFechaInicio { get; set; }
        public string VinculacionFechaFin { get; set; }
        public string Modalidad { get; set; }
        public int IsPernoctaUltimoDia { get; set; }
        public int IsTerminalAereo { get; set; }
        public int IdTerminales { get; set; }
        public string TerminalAereaNombre { get; set; }

        /// <summary>
        /// Origen de la fila cuando se audita una corrida del AG: NUEVA (staging) o CONSERVADA
        /// (horario aprobado verificado). Null en los flujos que no usan manifiesto de cobertura.
        /// No altera la persistencia de la sesión.
        /// </summary>
        public string OrigenSesion { get; set; }
        
        /// <summary>
        /// Valor del gasto de terminal aérea asignado a la vinculación.
        /// Proviene de la columna 'gastoTerminalAerea' en la tabla vinculaciones.
        /// </summary>
        public double GastoTerminalAerea { get; set; }
        
        /// <summary>
        /// Costo de transporte terrestre hacia esta sede.
        /// Proviene de la columna 'transporte' en la tabla de sedes (JOIN en vw_SesionesDetalle).
        /// </summary>
        public decimal? TransporteSede { get; set; }

        public string ProgramaDetalle => $"{ProgramaAbreviatura} ({ProgramaTipo})";

        public string RangoFechas => FechaSesion.HasValue
            ? FechaSesion.Value.ToString("dd/MM/yyyy")
            : "No definida";
    }
}

