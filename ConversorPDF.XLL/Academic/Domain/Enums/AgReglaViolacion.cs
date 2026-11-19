namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Código de regla o restricción evaluada por el motor genético.
    /// Hard constraints (R1-R7, W1) penalizan con infactibilidad.
    /// Soft constraints (S1-S5) penalizan la aptitud relativa.
    /// </summary>
    public enum AgReglaViolacion
    {
        // Restricciones Duras Canónicas del Plan (Hard Constraints)
        R1_SuperposicionDocente = 1,
        R2_SuperposicionAula = 2,
        R3_SuperposicionGrupo = 3,
        R4_DosAsignaturasMismoDia = 4,          // Salvo sábado (2x4h o 1x8h)
        R5_MaxAsignaturasSemana = 5,            // Máx 2 asignaturas por grupo/semana
        R6_SecuenciaAsignaturas = 6,            // Finalizar antes de iniciar nuevas
        R7_TerceroOtroGrupoFechaHora = 7,       // Misma fecha+hora en otro grupo
        W1_VentanaSemestralFueraDeRango = 8,
        W1_VentanaSemestre = 8,

        // Restricciones Duras Adicionales
        R4_CompatibilidadAula = 20,
        R4b_CompatibilidadAula = 20,
        R5_HabilitacionDocentePool = 21,
        R5b_HabilitacionDocentePool = 21,
        R6_DisponibilidadDocente = 22,
        R6b_DisponibilidadDocente = 22,
        R7_CoberturaHorariaIncompleta = 23,
        R7b_CoberturaHorariaIncompleta = 23,
        R8_DiasDuplicadosGrupo = 24,            // Fase 456: Días repetidos en el mismo grupo de una asignatura multi-día
        R8_DiaFueraDePatron = 25,               // Fase 456: Día no permitido en la especificación del patrón semanal
        R9_VacanteConDocenteReal = 26,          // Fase 2019 / Fase 2024: Gen vacante asignado a docente real (≠ 0 y ≠ centinela)
        R9_CentinelaEnNoVacante = 27,           // Fase 464: Gen no vacante asignado al tercero centinela VACANTE
        R10_CoherenciaDocenteGrupo = 28,        // Fase 544: Incoherencia de múltiples docentes en el mismo ClaveGrupo
        R11_DuracionSesionInvalida = 29,        // Fase 935 / Fase 1030: Reservada, no emitida; duración gobernada por bloques canónicos
        R12_CarrilesConcurrentesMismaAula = 30, // Fase 935 / Fase 1030: Reservada, no emitida; solape cubierto por R2/R14 sin riesgo de falsos positivos
        R13_OleadasDispersas = 31,              // Fase 935 / Fase 1030: Reservada, no emitida; oleadas diferidas a fase de promoción
        R14_BloqueHorarioNoAsignado = 32,       // Fase 900: Gen con carga horaria pero sin bloque asignado (0)
        R14_BloqueSolapadoPorIntervalo = 33,    // Fase 935/957: Bloques con solape temporal por intervalo matemático
        R15_VacanteExcedida = 34,               // Fase 2018: Total de vacantes forzadas supera el déficit estructural calculado
        R16_VacanteEstructuralNoConsentida = 35, // Fase 2054: Asignatura con vacante estructural sin consentimiento previo

        // Restricciones Blandas (Soft Constraints)
        S1_HuecosHorarios = 101,
        S2_DesplazamientoInterSede = 102,
        S3_PreferenciaFranjaDocente = 103,
        S4_SobrecargaDiaria = 104,
        S5_OptimizacionViaticos = 105,
        S6_HuecoSemanaCarril = 106,             // Fase 935 / Fase 1030: Reservada, no emitida
        S7_DispersorOleadas = 107               // Fase 935 / Fase 1030: Reservada, no emitida
    }
}
