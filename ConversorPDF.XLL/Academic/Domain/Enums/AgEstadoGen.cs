namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Estado de asignación y consistencia de un gen (unidad de programación curso-docente-bloque-aula).
    /// </summary>
    public enum AgEstadoGen
    {
        Valido = 0,
        Conflicto = 1,
        Conflictivo = 1,
        Incompleto = 2,
        Fijado = 3
    }
}
