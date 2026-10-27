namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Motivo por el cual una asignatura queda como vacante (sin docente asignable) durante la optimización (Fases 318, 329).
    /// </summary>
    public enum AgMotivoVacante
    {
        Ninguno = 0,
        SIN_CANDIDATOS = 1,
        SIN_DISPONIBILIDAD = 2,
        NO_CUMPLE_REQUISITOS = 3,
        POOL_VACIO = 4,
        FUERA_DE_VENTANA = 5,
        SIN_PLAN_OLEADA = 6,
        SIN_BLOQUE_NIVEL = 7
    }
}
