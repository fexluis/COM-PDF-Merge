using System;
using System.Collections.Generic;
using System.Linq;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Registro histórico de ejecución de una corrida del motor genético.
    /// Mapea a la tabla 'ag_corrida'.
    /// </summary>
    public class AgCorrida
    {
        public long Id { get; set; }
        public int IdPeriodo { get; set; }
        public int IdPrograma { get; set; }
        public int IdAula { get; set; }
        public string FechaInicioCorrida { get; set; }
        public string FechaFinCorrida { get; set; }
        public AgEstadoCorrida Estado { get; set; }
        public int PoblacionTamano { get; set; }
        public int GeneracionesMax { get; set; }
        public int GeneracionesEjecutadas { get; set; }
        public double TasaCruce { get; set; }
        public double TasaMutacion { get; set; }
        public int Elitismo { get; set; }
        public double MejorFitness { get; set; }
        public int HardViolations { get; set; }
        public double SoftViolations { get; set; }
        public int Vacantes { get; set; }
        public long? IdIndividuoOptimo { get; set; }
        public string ParametrosJson { get; set; }
        public string Subunidad { get; set; }

        private HashSet<int> _vacantesConsentidas = new HashSet<int>();

        /// <summary>
        /// FASE 1883/1884/1885: Asignaturas autorizadas para docente vacante en esta corrida.
        /// </summary>
        public HashSet<int> VacantesConsentidas
        {
            get => _vacantesConsentidas ?? (_vacantesConsentidas = new HashSet<int>());
            set => _vacantesConsentidas = value ?? new HashSet<int>();
        }

        /// <summary>
        /// FASE 1885/1887: Columna tipada 'vacantes_consentidas' en ag_corrida ("101,102") para mapeo automático Dapper.
        /// Se ignora en cualquier serialización para evitar duplicar el estado del dominio.
        /// </summary>
        [System.Runtime.Serialization.IgnoreDataMember]
        [Newtonsoft.Json.JsonIgnore]
        public string vacantes_consentidas
        {
            get => string.Join(",", VacantesConsentidas.OrderBy(x => x));
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    _vacantesConsentidas = new HashSet<int>();
                    return;
                }
                _vacantesConsentidas = new HashSet<int>(
                    value.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                         .Select(s => int.TryParse(s.Trim(), out var v) ? v : 0)
                         .Where(v => v > 0));
            }
        }

        public AgCorrida()
        {
            FechaInicioCorrida = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Estado = AgEstadoCorrida.Iniciada;
            PoblacionTamano = 50;
            GeneracionesMax = 200;
            TasaCruce = 0.85;
            TasaMutacion = 0.10;
            Elitismo = 2;
            ParametrosJson = "{}";
            Subunidad = string.Empty;
        }
    }
}
