using System;
using ConversorPDF.XLL.Academic.Domain.Enums;

namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    /// <summary>
    /// Ponderaciones y coeficientes de penalización configurables para la función de aptitud (fitness) (Fase 948).
    /// </summary>
    public class AgPesosFitness
    {
        // Penalizaciones Duras (Hard Penalty Base)
        public double PenalidadHardBase { get; set; } = 1000.0;
        public double PenalidadHardSuperposicionDocente { get; set; } = 1000.0;
        public double PenalidadHardSuperposicionAula { get; set; } = 1000.0;
        public double PenalidadHardSuperposicionGrupo { get; set; } = 1000.0;
        public double PenalidadHardDosAsignaturasDia { get; set; } = 900.0;
        public double PenalidadHardMaxAsignaturasSemana { get; set; } = 800.0;
        public double PenalidadHardSecuencia { get; set; } = 850.0;
        public double PenalidadHardTerceroOtroGrupo { get; set; } = 1000.0;
        public double PenalidadHardCompatibilidadAula { get; set; } = 800.0;
        public double PenalidadHardPoolDocente { get; set; } = 1000.0;
        public double PenalidadHardDisponibilidadDocente { get; set; } = 900.0;
        public double PenalidadHardCoberturaHoraria { get; set; } = 1000.0;
        public double PenalidadHardVentanaSemestral { get; set; } = 1000.0;
        public double PenalidadHardR11_Duracion { get; set; } = 1000.0;     // Fase 935 / Fase 1030: reservada
        public double PenalidadHardR12_CarrilesAula { get; set; } = 1000.0; // Fase 935 / Fase 1030: reservada
        public double PenalidadHardR13_Oleadas { get; set; } = 1000.0;      // Fase 935 / Fase 1030: reservada
        public double PenalidadHardR14_Bloque { get; set; } = 1000.0;

        // Ponderaciones Blandas (Soft Weights)
        public double PesoS1_HuecosHorarios { get; set; } = 10.0;
        public double PesoS2_DesplazamientoInterSede { get; set; } = 25.0;
        public double PesoS3_PreferenciaFranja { get; set; } = 5.0;
        public double PesoS4_SobrecargaDiaria { get; set; } = 15.0;
        public double PesoS5_OptimizacionViaticos { get; set; } = 20.0;
        public double PesoS6_HuecoSemanaCarril { get; set; } = 15.0;        // Fase 935 / Fase 1030: reservada
        public double PesoS7_DispersorOleadas { get; set; } = 20.0;         // Fase 935 / Fase 1030: reservada
        public double PesoVacante { get; set; } = 50.0;

        public double PorRegla(AgReglaViolacion regla)
        {
            switch (regla)
            {
                case AgReglaViolacion.R1_SuperposicionDocente: return PenalidadHardSuperposicionDocente;
                case AgReglaViolacion.R2_SuperposicionAula: return PenalidadHardSuperposicionAula;
                case AgReglaViolacion.R3_SuperposicionGrupo: return PenalidadHardSuperposicionGrupo;
                case AgReglaViolacion.R4_DosAsignaturasMismoDia: return PenalidadHardDosAsignaturasDia;
                case AgReglaViolacion.R5_MaxAsignaturasSemana: return PenalidadHardMaxAsignaturasSemana;
                case AgReglaViolacion.R6_SecuenciaAsignaturas: return PenalidadHardSecuencia;
                case AgReglaViolacion.R7_TerceroOtroGrupoFechaHora: return PenalidadHardTerceroOtroGrupo;
                case AgReglaViolacion.W1_VentanaSemestralFueraDeRango: return PenalidadHardVentanaSemestral;
                case AgReglaViolacion.R4_CompatibilidadAula: return PenalidadHardCompatibilidadAula;
                case AgReglaViolacion.R5_HabilitacionDocentePool: return PenalidadHardPoolDocente;
                case AgReglaViolacion.R6_DisponibilidadDocente: return PenalidadHardDisponibilidadDocente;
                case AgReglaViolacion.R7_CoberturaHorariaIncompleta: return PenalidadHardCoberturaHoraria;
                case AgReglaViolacion.R8_DiasDuplicadosGrupo: return PenalidadHardBase;
                case AgReglaViolacion.R8_DiaFueraDePatron: return PenalidadHardBase;
                case AgReglaViolacion.R9_VacanteConDocenteReal: return PenalidadHardBase;
                case AgReglaViolacion.R9_CentinelaEnNoVacante: return PenalidadHardBase;
                case AgReglaViolacion.R10_CoherenciaDocenteGrupo: return PenalidadHardBase;
                case AgReglaViolacion.R11_DuracionSesionInvalida: return PenalidadHardR11_Duracion;
                case AgReglaViolacion.R12_CarrilesConcurrentesMismaAula: return PenalidadHardR12_CarrilesAula;
                case AgReglaViolacion.R13_OleadasDispersas: return PenalidadHardR13_Oleadas;
                case AgReglaViolacion.R14_BloqueHorarioNoAsignado: return PenalidadHardR14_Bloque;
                case AgReglaViolacion.R14_BloqueSolapadoPorIntervalo: return PenalidadHardR14_Bloque;
                case AgReglaViolacion.R15_VacanteExcedida: return PenalidadHardBase;
                case AgReglaViolacion.R16_VacanteEstructuralNoConsentida: return PenalidadHardBase;
                case AgReglaViolacion.S1_HuecosHorarios: return PesoS1_HuecosHorarios;
                case AgReglaViolacion.S2_DesplazamientoInterSede: return PesoS2_DesplazamientoInterSede;
                case AgReglaViolacion.S3_PreferenciaFranjaDocente: return PesoS3_PreferenciaFranja;
                case AgReglaViolacion.S4_SobrecargaDiaria: return PesoS4_SobrecargaDiaria;
                case AgReglaViolacion.S5_OptimizacionViaticos: return PesoS5_OptimizacionViaticos;
                case AgReglaViolacion.S6_HuecoSemanaCarril: return PesoS6_HuecoSemanaCarril;
                case AgReglaViolacion.S7_DispersorOleadas: return PesoS7_DispersorOleadas;
                default: return PenalidadHardBase;
            }
        }

        public static AgPesosFitness PorDefecto()
        {
            return new AgPesosFitness();
        }
    }
}
