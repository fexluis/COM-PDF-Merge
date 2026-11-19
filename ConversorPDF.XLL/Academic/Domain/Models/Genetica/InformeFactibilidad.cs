using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Informe de diagnóstico y factibilidad pre-vuelo (pre-flight) emitido antes de iniciar la corrida del AG.
    /// Valida si matemáticamente la ventana y los recursos satisfacen la demanda académica.
    /// </summary>
    public class InformeFactibilidad
    {
        public bool EsFactible { get; set; }
        public double TotalHorasRequeridas { get; set; }
        public double TotalCapacidadHorariaAulas { get; set; }
        public int TotalCursosProgramar { get; set; }
        public int TotalDocentesDisponibles { get; set; }
        public List<string> BloqueosCriticos { get; set; }
        public List<string> Bloqueos => BloqueosCriticos;
        public List<string> Advertencias { get; set; }
        public Dictionary<int, AgMotivoVacante> AsignaturasVacantes { get; set; }
        public Dictionary<int, int> DeficitDocentePorAsignatura { get; set; }

        public InformeFactibilidad()
        {
            EsFactible = true;
            BloqueosCriticos = new List<string>();
            Advertencias = new List<string>();
            AsignaturasVacantes = new Dictionary<int, AgMotivoVacante>();
            DeficitDocentePorAsignatura = new Dictionary<int, int>();
        }

        public void AgregarBloqueo(string mensaje)
        {
            EsFactible = false;
            BloqueosCriticos.Add(mensaje);
        }

        public void AgregarAdvertencia(string mensaje)
        {
            Advertencias.Add(mensaje);
        }
    }
}
