using System;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    /// <summary>
    /// Representa una vinculación académica que asocia un docente (Tercero) con una asignatura,
    /// sede, programa y semestre, definiendo el cronograma de horas semanal.
    /// </summary>
    public class Vinculacion
    {
        public int Id { get; set; }
        public int IdPeriodo { get; set; }
        public int IdSede { get; set; }

        /// <summary>FK al docente/contratista responsable de esta vinculación.</summary>
        public int IdTercero { get; set; }

        public int IdPrograma { get; set; }
        
        /// <summary>
        /// Semestre: I, II, III, IV, V, VI, VII, VIII, IX, X
        /// </summary>
        public string Semestre { get; set; }
        
        public int IdAsignatura { get; set; }
        public string Grupo { get; set; }
        public int TotalHoras { get; set; }
        
        /// <summary>
        /// Almacenado como cadena separada por comas. Ej: "JUEVES|8, VIERNES|4"
        /// </summary>
        public string DiaHorasSemanales { get; set; }
        
        public DateTime? FechaInicio { get; set; }
        public DateTime? FechaFin { get; set; }
        public string Modalidad { get; set; }
        
        public int IsPernoctaUltimoDia { get; set; }
        public int IsTerminalAereo { get; set; }
        public int IdTerminales { get; set; }
        public double GastoTerminalAerea { get; set; }
        public int EsVacante { get; set; }
        public string MotivoVacante { get; set; } = string.Empty;
        public string Subunidad { get; set; }
    }
}
