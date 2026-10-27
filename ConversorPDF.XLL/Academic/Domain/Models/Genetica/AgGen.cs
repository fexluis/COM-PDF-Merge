using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Representa un gen en el cromosoma: asignación de (Docente, Aula, Bloque) a una unidad curricular.
    /// Mapea a la tabla 'ag_gen'.
    /// </summary>
    public class AgGen
    {
        public long Id { get; set; }
        public long IdIndividuo { get; set; }
        public int IdPeriodo { get; set; }
        public int IdAsignatura { get; set; }
        public int IdPrograma { get; set; }
        public int IdSede { get; set; }
        public string Semestre { get; set; }
        public string Grupo { get; set; }
        public int IdTercero { get; set; }
        public int IdAula { get; set; }
        public int IdBloqueHorario { get; set; }
        public int DiaSemana { get; set; } = 1;
        public double HorasSemanales { get; set; }
        public double TotalHorasAsignatura { get; set; }
        public DateTime? FechaInicioEstimada { get; set; }
        public DateTime? FechaFinEstimada { get; set; }
        private int _esFijo;
        public int EsFijo
        {
            get => _esFijo;
            set
            {
                _esFijo = value;
                if (_esFijo == 1)
                {
                    Estado = AgEstadoGen.Fijado;
                }
            }
        }
        public int EsVacante { get; set; }
        public string MotivoVacante { get; set; }
        public int CandidatosPool { get; set; }
        public int DentroDeVentana { get; set; } = 1;
        public AgEstadoGen Estado { get; set; }
        public string Subunidad { get; set; }
        /// <summary>
        /// Indica si la franja, día y horas del gen provienen de una decisión explícita de aula (β/γ/horas)
        /// y deben preservarse inmunes a mutación de aula/bloque y cruce de horario, permitiendo optimizar el docente.
        /// </summary>
        public bool EsFijoTemporal { get; set; }

        // Propiedades de patrón semanal y multi-día (Fase 382 / Fase 397 / Fase 398)
        public int IndiceDiaPatron { get; set; }
        public string PatronSemanal { get; set; } = string.Empty;
        public string DiasPatron { get; set; } = string.Empty;
        public string ClaveGrupo { get; set; } = string.Empty;
        private double _horasGen = 4.0;
        public double HorasGen
        {
            get => _horasGen;
            set => _horasGen = value < 0.0 ? 0.0 : value;   // Fase 583
        }
        /// <summary>
        /// Fechas válidas específicas asignadas a este gen por el planificador semanal (Fase 414).
        /// </summary>
        public List<DateTime> FechasGen { get; set; } = new List<DateTime>();

        public AgGen()
        {
            Semestre = string.Empty;
            Grupo = "A";
            MotivoVacante = string.Empty;
            Estado = AgEstadoGen.Valido;
            Subunidad = string.Empty;
            PatronSemanal = string.Empty;
            DiasPatron = string.Empty;
            ClaveGrupo = string.Empty;
            HorasGen = 4.0;
            FechasGen = new List<DateTime>();
        }

        /// <summary>
        /// Sesiones que este gen (un día del patrón) aporta según el histograma semanal (Fase 382).
        /// </summary>
        public int SesionesDelGen()
        {
            if (string.IsNullOrWhiteSpace(PatronSemanal))
            {
                return (int)Math.Ceiling(TotalHorasAsignatura / Math.Max(1.0, HorasSemanales));
            }

            var partes = PatronSemanal.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            int count = 0;
            for (int i = 0; i < partes.Length; i++)
            {
                if (int.TryParse(partes[i], out int v) && IndiceDiaPatron < v)
                {
                    count++;
                }
            }
            return count;
        }

        /// <summary>
        /// Asigna un bloque horario manteniendo la coherencia de IdBloqueHorario, DiaSemana y HorasSemanales.
        /// </summary>
        public void AsignarBloque(AgBloqueHorario bloque)
        {
            if (bloque == null) throw new ArgumentNullException(nameof(bloque));
            IdBloqueHorario = bloque.Id;
            DiaSemana = bloque.DiaSemanaNumero;
            HorasSemanales = bloque.DuracionHoras;
        }

        /// <summary>
        /// Asigna un aula manteniendo la coherencia invariante de IdAula, IdPeriodo, IdSede, IdPrograma, Semestre y Grupo (Fase 159/162).
        /// </summary>
        public void AsignarAula(AgAula aula)
        {
            if (aula == null) throw new ArgumentNullException(nameof(aula));
            IdAula = aula.Id;
            IdPeriodo = aula.IdPeriodo;
            IdSede = aula.IdSede;
            IdPrograma = aula.IdPrograma;
            Semestre = aula.Semestre ?? string.Empty;
            Grupo = string.IsNullOrWhiteSpace(aula.Grupo) ? "A" : aula.Grupo;
        }

        public AgGen Clonar()
        {
            return new AgGen
            {
                Id = 0,
                IdIndividuo = this.IdIndividuo,
                IdPeriodo = this.IdPeriodo,
                IdAsignatura = this.IdAsignatura,
                IdPrograma = this.IdPrograma,
                IdSede = this.IdSede,
                Semestre = this.Semestre,
                Grupo = this.Grupo,
                IdTercero = this.IdTercero,
                IdAula = this.IdAula,
                IdBloqueHorario = this.IdBloqueHorario,
                DiaSemana = this.DiaSemana,
                HorasSemanales = this.HorasSemanales,
                TotalHorasAsignatura = this.TotalHorasAsignatura,
                FechaInicioEstimada = this.FechaInicioEstimada,
                FechaFinEstimada = this.FechaFinEstimada,
                EsFijo = this.EsFijo,
                EsFijoTemporal = this.EsFijoTemporal,
                EsVacante = this.EsVacante,
                MotivoVacante = this.MotivoVacante,
                CandidatosPool = this.CandidatosPool,
                DentroDeVentana = this.DentroDeVentana,
                Estado = this.Estado,
                Subunidad = this.Subunidad,
                IndiceDiaPatron = this.IndiceDiaPatron,
                PatronSemanal = this.PatronSemanal,
                DiasPatron = this.DiasPatron,
                ClaveGrupo = this.ClaveGrupo,
                HorasGen = this.HorasGen,
                FechasGen = this.FechasGen != null ? new List<DateTime>(this.FechasGen) : new List<DateTime>()
            };
        }

        public override string ToString()
        {
            return $"Gen[Asig:{IdAsignatura}, Prog:{IdPrograma}, Doc:{IdTercero}, Aula:{IdAula}, Bloque:{IdBloqueHorario}, Sem:{Semestre}]";
        }
    }
}
